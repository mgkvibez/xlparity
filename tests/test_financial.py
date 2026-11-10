"""Cohort 4: the financial functions, where a plausible answer and a correct one look alike.

Every case here is a decision someone made on purpose. NPV starting at period 1, DB rounding its rate to three
decimals before depreciating anything, DDB stopping at salvage rather than at the end of its formula, and IRR
refusing a cash flow that never changes sign instead of returning a rate that happens to be a root of nothing.
"""

import pytest

from xlparity import ExcelError, evaluate, parse


def value(source, cells=None):
    return evaluate(parse(source), cells or {})


def number(source, cells=None):
    result = value(source, cells)
    assert isinstance(result, (int, float)) and not isinstance(result, bool), result
    return float(result)


def error(source, cells=None):
    result = value(source, cells)
    assert isinstance(result, ExcelError), result
    return result.code


# NPV ---------------------------------------------------------------------


def test_npv_discounts_the_first_value_by_one_period():
    # 100/1.05 + 200/1.05^2 + 300/1.05^3; an engine starting at period 0 would return 635.24 instead
    assert number("NPV(0.05,100,200,300)") == pytest.approx(535.7952704891479)


def test_npv_at_zero_rate_is_plain_summation():
    assert number("NPV(0,A1:A3)", {"A1": 100, "A2": 200, "A3": 300}) == 600


def test_npv_reads_a_range_and_skips_text_inside_it():
    cells = {"A1": 1000, "A2": "n/a", "A3": 1000}
    # the text cell is not discounted as a period of its own, so the second 1000 is period 2
    assert number("NPV(0.1,A1:A3)", cells) == pytest.approx(1000 / 1.1 + 1000 / 1.1**2)


def test_npv_propagates_an_error_argument():
    assert error('NPV(0.1,"abc")') == "#VALUE!"


def test_npv_at_minus_one_hundred_percent_is_a_division_by_zero():
    assert error("NPV(-1,100,200)") == "#DIV/0!"


def test_npv_needs_a_rate_and_a_value():
    assert error("NPV(0.1)") == "#VALUE!"


# IRR ---------------------------------------------------------------------


def test_irr_of_a_doubling_investment_is_ten_percent():
    assert number("IRR(A1:A2)", {"A1": -1000, "A2": 1100}) == pytest.approx(0.1, abs=1e-12)


def test_irr_solves_a_four_year_return():
    cells = {"A1": -1000, "A2": 300, "A3": 300, "A4": 300, "A5": 300}
    rate = number("IRR(A1:A5)", cells)
    # the definition, checked as a residual: NPV at the returned rate is zero
    residual = sum(amount / (1 + rate) ** period for period, amount in enumerate([-1000, 300, 300, 300, 300]))
    assert abs(residual) < 1e-9 * 1000


def test_irr_uses_the_guess_as_a_starting_point_not_as_an_answer():
    cells = {"A1": -1000, "A2": 300, "A3": 300, "A4": 300, "A5": 300}
    assert number("IRR(A1:A5,0.5)", cells) == pytest.approx(number("IRR(A1:A5)", cells), abs=1e-9)


def test_irr_refuses_a_cash_flow_that_never_changes_sign():
    assert error("IRR(A1:A2)", {"A1": 100, "A2": 200}) == "#NUM!"


def test_irr_refuses_a_guess_at_or_below_minus_one():
    cells = {"A1": -1000, "A2": 1100}
    assert error("IRR(A1:A2,-1)", cells) == "#NUM!"


def test_irr_takes_an_argument_count_it_can_work_with():
    assert error("IRR()") == "#VALUE!"
    assert error("IRR(A1:A2,0.1,0.2)", {"A1": -1000, "A2": 1100}) == "#VALUE!"


# XIRR --------------------------------------------------------------------


def test_xirr_uses_actual_days_over_three_sixty_five():
    cells = {"A1": -1000, "A2": 500, "A3": 700, "B1": 1, "B2": 183, "B3": 366}
    rate = number("XIRR(A1:A3,B1:B3)", cells)
    # 182 days is 0.49863 years, and the residual is what proves the day count
    residual = 500 / (1 + rate) ** (182 / 365) + 700 / (1 + rate) ** (365 / 365) - 1000
    assert abs(residual) < 1e-9 * 1000


def test_xirr_refuses_mismatched_lengths():
    cells = {"A1": -1000, "A2": 500, "B1": 1}
    assert error("XIRR(A1:A2,B1:B1)", cells) == "#NUM!"


def test_xirr_refuses_two_dates_for_one_cash_flow():
    cells = {"A1": -1000, "B1": 1, "B2": 366}
    assert error("XIRR(A1:A1,B1:B2)", cells) == "#NUM!"


def test_xirr_refuses_a_series_without_a_sign_change():
    cells = {"A1": 100, "A2": 200, "B1": 1, "B2": 100}
    assert error("XIRR(A1:A2,B1:B2)", cells) == "#NUM!"


def test_xirr_ignores_text_in_the_value_column():
    cells = {"A1": -1000, "A2": "n/a", "A3": 1100, "B1": 1, "B2": 2, "B3": 366}
    rate = number("XIRR(A1:A3,B1:B3)", cells)
    assert rate == pytest.approx(0.1, abs=1e-3)  # one year, ten percent, to the day count's precision


# SLN, SYD ----------------------------------------------------------------


def test_sln_spreads_the_depreciable_amount_evenly():
    assert number("SLN(30000,7500,10)") == 2250


def test_sln_refuses_a_zero_life_with_a_division_error():
    assert error("SLN(1000,100,0)") == "#DIV/0!"


def test_syd_front_loads_the_depreciation():
    assert number("SYD(30000,7500,10,1)") == pytest.approx(4090.909090909091)
    assert number("SYD(30000,7500,10,10)") == pytest.approx(409.0909090909091)


def test_syd_refuses_a_period_past_the_life():
    assert error("SYD(30000,7500,10,11)") == "#NUM!"


# DB, DDB -----------------------------------------------------------------


def test_db_rounds_its_rate_to_three_decimals_before_depreciating():
    # 1 - (0.1)^(1/6) is 0.318713..., so the rate is 0.319 and year one is 319000, not 318713.35
    assert number("DB(1000000,100000,6,1)") == 319000


def test_db_prorates_its_first_year_by_the_month_argument():
    assert number("DB(1000000,100000,6,1,6)") == pytest.approx(159500)
    assert number("DB(1000000,100000,6,2,6)") == pytest.approx((1000000 - 159500) * 0.319)


def test_db_adds_a_partial_year_after_the_life_ends():
    cells = number("DB(1000000,100000,6,6,6)")
    total = sum(number(f"DB(1000000,100000,6,{period},6)") for period in range(1, 7))
    assert number("DB(1000000,100000,6,7,6)") == pytest.approx((1000000 - total) * 0.319 * 6 / 12)
    assert cells > 0


def test_db_refuses_arguments_that_cannot_describe_an_asset():
    assert error("DB(1000,100,0,1)") == "#NUM!"
    assert error("DB(1000,100,5,1,13)") == "#NUM!"
    assert error("DB(1000,100,5,7)") == "#NUM!"


def test_ddb_doubles_the_straight_line_rate():
    assert number("DDB(2400,300,10,1)") == 480
    assert number("DDB(2400,300,10,2)") == 384


def test_ddb_stops_at_salvage_instead_of_following_its_formula():
    # year ten would be 64.42 by the formula; only 22.12 is left above salvage
    assert number("DDB(2400,300,10,10)") == pytest.approx(22.1225472)


def test_ddb_honours_a_factor_other_than_two():
    assert number("DDB(2400,300,10,3,1.5)") == pytest.approx(2400 * 0.85**2 * 0.15)


def test_ddb_refuses_a_period_past_the_life_and_a_useless_factor():
    assert error("DDB(2400,300,10,11)") == "#NUM!"
    assert error("DDB(2400,300,10,1,0)") == "#NUM!"


def test_financial_functions_propagate_an_error_rather_than_guessing():
    assert error("SLN(1000,100,SQRT(-1))") == "#NUM!"
    assert error("DDB(2400,300,10,MOD(1,0))") == "#DIV/0!"
    assert error("NPV(0.1,1*SQRT(-1))") == "#NUM!"
    assert error("SYD(30000,7500,10)") == "#VALUE!"
