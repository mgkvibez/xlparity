"""Dynamic arrays, as far as one formula can be true.

The half of dynamic arrays that lives inside a formula is a value: SEQUENCE(5) is five numbers, and SUM of
them is 15 whether or not anything spills. The other half is a place on the sheet, and this engine evaluates a
formula in one cell, so a result with cells of its own is #SPILL! rather than a quietly truncated first value.

Every expectation below is arithmetic that can be checked by hand, not a number read out of a spreadsheet this
repository has not seen.
"""


from xlparity import evaluate, parse


def run(formula, cells=None):
    return evaluate(parse(formula), cells or {})


def code(value):
    return getattr(value, "code", value)


CELLS = {"A1": 3, "A2": 1, "A3": 2, "A4": 2, "B1": 1, "B2": 0, "B3": 1, "B4": 3}


class TestSequence:
    def test_counts_from_one_by_default(self):
        assert run("SUM(SEQUENCE(5))") == 15
        assert run("SUM(SEQUENCE(10))") == 55

    def test_fills_rows_before_columns(self):
        # a 2 by 3 block is 1 2 3 over 4 5 6, so its last cell is 6 at row 2 column 3
        assert run("SUM(SEQUENCE(2,3))") == 21
        assert run("INDEX(SEQUENCE(2,3),2,3)") == 6
        assert run("INDEX(SEQUENCE(2,3),1,2)") == 2

    def test_takes_a_start_and_a_step(self):
        assert run("SUM(SEQUENCE(3,1,0,2))") == 6  # 0 2 4
        assert run("SUM(SEQUENCE(4,1,1,-1))") == -2  # 1 0 -1 -2
        assert run("INDEX(SEQUENCE(3,1,10,5),3)") == 20

    def test_an_empty_array_is_calc_rather_than_zero_rows(self):
        assert code(run("SEQUENCE(0)")) == "#CALC!"
        assert code(run("SEQUENCE(3,0)")) == "#CALC!"

    def test_refuses_arguments_of_the_wrong_shape(self):
        assert code(run("SEQUENCE()")) == "#VALUE!"
        assert code(run("SEQUENCE(2,2,1,1,1)")) == "#VALUE!"
        assert code(run('SEQUENCE("two")')) == "#VALUE!"


class TestUnique:
    def test_keeps_first_appearance_order_rather_than_sorting(self):
        cells = {"A1": 3, "A2": 1, "A3": 3, "A4": 2}
        # 3, 1, 2 in that order: UNIQUE sorts nothing, so the middle value is what proves the order
        assert run("SUM(UNIQUE(A1:A4))", cells) == 6
        assert run("INDEX(UNIQUE(A1:A4),1)", cells) == 3
        assert run("INDEX(UNIQUE(A1:A4),3)", cells) == 2

    def test_collapses_text_as_well_as_numbers(self):
        cells = {"A1": "b", "A2": "a", "A3": "b"}
        assert run("COUNTA(UNIQUE(A1:A4))", cells) == 2

    def test_a_scalar_is_not_an_array(self):
        assert code(run("UNIQUE(3)")) == "#VALUE!"

    def test_de_duplicating_columns_is_refused_rather_than_treated_as_rows(self):
        assert code(run("UNIQUE(A1:A4,TRUE)", CELLS)) == "#VALUE!"


class TestSort:
    def test_sorts_ascending_by_default_and_keeps_every_value(self):
        assert run("SUM(SORT(A1:A4))", CELLS) == 8
        assert run("INDEX(SORT(A1:A4),1)", CELLS) == 1
        assert run("INDEX(SORT(A1:A4),4)", CELLS) == 3

    def test_sorts_descending_when_told_to(self):
        assert run("INDEX(SORT(A1:A4,1,-1),1)", CELLS) == 3
        assert run("INDEX(SORT(A1:A4,1,-1),4)", CELLS) == 1

    def test_orders_numbers_before_text_the_way_excel_does(self):
        cells = {"A1": "b", "A2": 2, "A3": "a", "A4": 1}
        assert run("INDEX(SORT(A1:A4),1)", cells) == 1
        assert run("INDEX(SORT(A1:A4),3)", cells) == "a"
        assert run("INDEX(SORT(A1:A4),4)", cells) == "b"

    def test_refuses_a_sort_index_a_vector_does_not_have(self):
        assert code(run("SORT(A1:A4,2)", CELLS)) == "#VALUE!"
        assert code(run("SORT(A1:A4,1,0)", CELLS)) == "#VALUE!"


class TestFilter:
    def test_keeps_the_values_whose_flag_is_true(self):
        # flags are 1, 0, 1, 3 so A1 and A3 and A4 are kept: 3 + 2 + 2
        assert run("SUM(FILTER(A1:A4,B1:B4))", CELLS) == 7
        assert run("INDEX(FILTER(A1:A4,B1:B4),2)", CELLS) == 2

    def test_a_flag_column_can_hold_the_words_true_and_false(self):
        cells = {"A1": 1, "A2": 2, "A3": 3, "B1": "TRUE", "B2": "FALSE", "B3": "TRUE"}
        assert run("SUM(FILTER(A1:A3,B1:B3))", cells) == 4

    def test_nothing_kept_is_calc_unless_an_if_empty_is_given(self):
        cells = {"A1": 1, "A2": 2, "B1": 0, "B2": 0}
        assert code(run("FILTER(A1:A2,B1:B2)", cells)) == "#CALC!"
        assert run('FILTER(A1:A2,B1:B2,"none")', cells) == "none"

    def test_the_flags_have_to_match_the_values_one_for_one(self):
        assert code(run("FILTER(A1:A4,B1:B2)", CELLS)) == "#VALUE!"


class TestSpilling:
    def test_an_array_result_is_a_spill_and_says_so(self):
        # the result has cells to occupy and this engine evaluates one formula in one cell
        assert code(run("SEQUENCE(3)")) == "#SPILL!"
        assert code(run("SORT(A1:A4)", CELLS)) == "#SPILL!"
        assert code(run("FILTER(A1:A4,B1:B4)", CELLS)) == "#SPILL!"

    def test_a_bare_range_is_still_value_because_it_is_not_an_array_that_was_computed(self):
        assert code(run("A1:A3", CELLS)) == "#VALUE!"

    def test_an_array_can_be_addressed_whole_now_that_array_values_exist(self):
        assert run("SUM(INDEX(A1:B4,0,1))", {"A1": 1, "B1": 9, "A2": 2, "B2": 8}) == 3
        assert run("SUM(INDEX(A1:B1,1,0))", {"A1": 1, "B1": 9}) == 10
