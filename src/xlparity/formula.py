"""A spreadsheet formula parser and evaluator, starting from the parts Excel is strict about.

The reason this project exists is parity, not a roughly-right calculator. Parity starts with the details
that "Excel-ish" engines get wrong: operator precedence that matches Excel exactly, comparison rules,
error values propagating as values instead of raising, and the difference between a blank cell and an
explicit zero.

Only a handful of functions are implemented so far. The parser is the foundation; the scoreboard
grows function by function, each measured against Excel's own output rather than against a guess.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from typing import Callable, Dict, List, Sequence, Tuple, Union

Number = Union[int, float]
Value = Union[Number, str, bool, None, "ExcelError"]


@dataclass(frozen=True)
class ExcelError:
    """An error value: data in Excel. It propagates, it compares, and ISERROR can test it."""

    code: str

    def __str__(self) -> str:  # pragma: no cover - display only
        return self.code


DIV0 = ExcelError("#DIV/0!")
VALUE = ExcelError("#VALUE!")
NAME = ExcelError("#NAME?")
REF = ExcelError("#REF!")
NUM = ExcelError("#NUM!")


class FormulaError(Exception):
    """Raised for syntax problems only. Excel error *values* are returned, never raised."""


# ---------------------------------------------------------------------------
# Tokenizer
# ---------------------------------------------------------------------------

Token = tuple[str, object]

_TOKEN_SPECIAL = set("+-*/^&%,()<>= ")


def tokenize(source: str) -> List[Token]:
    tokens: List[Token] = []
    i = 0
    length = len(source)

    while i < length:
        char = source[i]

        if char.isspace():
            i += 1
            continue

        if char == '"':
            j = i + 1
            text = []
            while j < length:
                if source[j] == '"':
                    if j + 1 < length and source[j + 1] == '"':  # doubled quote is a literal quote
                        text.append('"')
                        j += 2
                        continue
                    break
                text.append(source[j])
                j += 1
            if j >= length:
                raise FormulaError("unterminated string literal")
            tokens.append(("string", "".join(text)))
            i = j + 1
            continue

        # Two-character comparison operators first, so <= is not read as < then =.
        if source[i : i + 2] in ("<=", ">=", "<>"):
            tokens.append(("op", source[i : i + 2]))
            i += 2
            continue

        if char in "+-*/^&%,()<>=.:":
            if char == "." and i + 1 < length and source[i + 1].isdigit():
                j = i
                while j < length and (source[j].isdigit() or source[j] == "."):
                    j += 1
                tokens.append(("number", float(source[i:j])))
                i = j
                continue
            if char in "(),":
                tokens.append((char, char))
            else:
                tokens.append(("op", char))
            i += 1
            continue

        if char.isdigit():
            j = i
            while j < length and (source[j].isdigit() or source[j] == "."):
                j += 1
            if j < length and source[j] in "eE":
                k = j + 1
                if k < length and source[k] in "+-":
                    k += 1
                while k < length and source[k].isdigit():
                    k += 1
                if k > j + 1:
                    j = k
            tokens.append(("number", float(source[i:j])))
            i = j
            continue

        if char.isalpha() or char == "$" or char == "_":
            j = i
            while j < length and (source[j].isalnum() or source[j] in "$._"):
                j += 1
            word = source[i:j]
            upper = word.upper()
            if upper in ("TRUE", "FALSE"):
                tokens.append(("bool", upper == "TRUE"))
            elif j < length and source[j] == "(":
                tokens.append(("func", upper))
            else:
                tokens.append(("ref", upper))
            i = j
            continue

        raise FormulaError(f"unexpected character {char!r}")

    return tokens


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------
#
# Excel precedence, loosest to tightest: comparison, concat (&), addition, multiplication,
# exponent (right-associative), percent, unary minus, parens.

_ADDITIVE = {"+", "-"}
_MULTIPLICATIVE = {"*", "/"}
_COMPARISON = {"=", "<>", "<", ">", "<=", ">="}


@dataclass(frozen=True)
class Literal:
    value: object


@dataclass(frozen=True)
class Range:
    """A rectangular block of cells, such as A1:C3.

    Excel treats a block differently from a scalar in one way that matters for parity: SUM(A1:A3) ignores text
    and logicals inside the block, while SUM("abc") is a #VALUE! error. Tagging values that came from a range
    is how the aggregates tell those two apart.
    """

    start: str
    end: str


class RangeValues(list):
    """Values read out of a Range, row-major, with the block's shape recorded.

    A list subclass, so every existing consumer keeps working; the shape is what lets VLOOKUP, INDEX and MATCH
    address a cell instead of a position.
    """

    rows: int = 0
    columns: int = 0

    def cell(self, row: int, column: int) -> "Value":
        """1-based, Excel-style access. Out of range is the caller's problem: they raise #REF!."""
        return self[(row - 1) * self.columns + (column - 1)]


@dataclass(frozen=True)
class Reference:
    name: str


@dataclass(frozen=True)
class Unary:
    op: str
    operand: "Node"


@dataclass(frozen=True)
class Binary:
    op: str
    left: "Node"
    right: "Node"


@dataclass(frozen=True)
class FunctionCall:
    name: str
    args: tuple


Node = Union[Literal, Reference, Unary, Binary, FunctionCall]


class _Parser:
    def __init__(self, tokens: Sequence[Token]) -> None:
        self._tokens = tokens
        self._index = 0

    def peek(self) -> Token | None:
        return self._tokens[self._index] if self._index < len(self._tokens) else None

    def next(self) -> Token:
        token = self.peek()
        if token is None:
            raise FormulaError("unexpected end of formula")
        self._index += 1
        return token

    def parse(self) -> Node:
        node = self.parse_comparison()
        if self.peek() is not None:
            raise FormulaError(f"unexpected token {self.peek()[1]!r}")
        return node

    def parse_comparison(self) -> Node:
        node = self.parse_concat()
        while True:
            token = self.peek()
            if token is None or token[0] != "op" or token[1] not in _COMPARISON:
                return node
            self.next()
            node = Binary(str(token[1]), node, self.parse_concat())

    def parse_concat(self) -> Node:
        node = self.parse_additive()
        while True:
            token = self.peek()
            if token is None or token[0] != "op" or token[1] != "&":
                return node
            self.next()
            node = Binary("&", node, self.parse_additive())

    def parse_additive(self) -> Node:
        node = self.parse_multiplicative()
        while True:
            token = self.peek()
            if token is None or token[0] != "op" or token[1] not in _ADDITIVE:
                return node
            self.next()
            node = Binary(str(token[1]), node, self.parse_multiplicative())

    def parse_multiplicative(self) -> Node:
        node = self.parse_exponent()
        while True:
            token = self.peek()
            if token is None or token[0] != "op" or token[1] not in _MULTIPLICATIVE:
                return node
            self.next()
            node = Binary(str(token[1]), node, self.parse_exponent())

    def parse_exponent(self) -> Node:
        node = self.parse_unary()
        token = self.peek()
        if token is not None and token[0] == "op" and token[1] == "^":
            self.next()
            # Right-associative: 2^3^2 is 2^(3^2) = 512, as in Excel.
            return Binary("^", node, self.parse_exponent())
        return node

    def parse_unary(self) -> Node:
        token = self.peek()
        if token is not None and token[0] == "op" and token[1] in ("+", "-"):
            self.next()
            return Unary(str(token[1]), self.parse_unary())
        return self.parse_postfix()

    def parse_postfix(self) -> Node:
        node = self.parse_primary()
        token = self.peek()
        if token is not None and token[0] == "op" and token[1] == "%":
            self.next()
            return Binary("/", node, Literal(100))
        return node

    def parse_primary(self) -> Node:
        token = self.next()
        kind, value = token

        if kind == "number":
            return Literal(value)
        if kind == "string":
            return Literal(value)
        if kind == "bool":
            return Literal(value)
        if kind == "ref":
            node: Node = Reference(str(value))
            following = self.peek()
            if following is not None and following[0] == "op" and following[1] == ":":
                self.next()
                closing = self.next()
                if closing[0] != "ref":
                    raise FormulaError("a range needs a cell reference on both sides of :")
                node = Range(str(value), str(closing[1]))
            return node
        if kind == "(":
            node = self.parse_comparison()
            closing = self.next()
            if closing[0] != ")":
                raise FormulaError("missing closing paren")
            return node
        if kind == "func":
            args: List[Node] = []
            opening = self.next()
            if opening[0] != "(":
                raise FormulaError("expected ( after function name")
            if self.peek() is not None and self.peek()[0] == ")":
                self.next()
                return FunctionCall(str(value), tuple())
            while True:
                args.append(self.parse_comparison())
                separator = self.next()
                if separator[0] == ")":
                    break
                if separator[0] != ",":
                    raise FormulaError("expected , or ) in argument list")
            return FunctionCall(str(value), tuple(args))

        raise FormulaError(f"unexpected token {value!r}")


def parse(source: str) -> Node:
    if not source.strip():
        raise FormulaError("empty formula")
    if source.lstrip().startswith("="):
        source = source.lstrip()[1:]
    return _Parser(tokenize(source)).parse()


# ---------------------------------------------------------------------------
# Evaluation
# ---------------------------------------------------------------------------


def _to_number(value: Value) -> Value:
    if isinstance(value, ExcelError):
        return value
    if value is None:
        return 0  # a blank cell is zero in arithmetic, but "" in concatenation
    if isinstance(value, bool):
        return 1 if value else 0
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip()
    try:
        return float(text)
    except ValueError:
        return VALUE


def _to_text(value: Value) -> str:
    if isinstance(value, ExcelError):
        return value.code
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _blank_as(other: Value) -> Value:
    """A blank takes the type of whatever it is compared against: 0, "" or FALSE."""
    if isinstance(other, bool):
        return False
    if isinstance(other, (int, float)):
        return 0.0
    return ""


def _rank(value: Value) -> int:
    """Excel sorts every type against every other: numbers first, then text, then logicals."""
    if isinstance(value, bool):
        return 2
    if isinstance(value, (int, float)):
        return 0
    return 1


def _compare(left: Value, right: Value) -> int:
    if left is None and right is None:
        return 0
    if left is None:
        return _compare(_blank_as(right), right)
    if right is None:
        return _compare(left, _blank_as(left))

    left_rank, right_rank = _rank(left), _rank(right)
    if left_rank != right_rank:
        return -1 if left_rank < right_rank else 1

    if left_rank == 0:
        a, b = float(left), float(right)  # type: ignore[arg-type]
        return (a > b) - (a < b)
    if left_rank == 2:
        return int(bool(left)) - int(bool(right))
    a, b = _to_text(left).lower(), _to_text(right).lower()
    return (a > b) - (a < b)


def _flatten(args: Sequence[Value]) -> List[Value]:
    flat: List[Value] = []
    for arg in args:
        if isinstance(arg, list):
            flat.extend(arg)
        else:
            flat.append(arg)
    return flat


def _fn_sum(args: Sequence[Value]) -> Value:
    total = 0.0
    for value in _flatten(args):
        if isinstance(value, ExcelError):
            return value
        if value is None or isinstance(value, str):  # SUM ignores text and blanks in ranges
            continue
        total += float(_to_number(value))
    return total


def _fn_round(args: Sequence[Value]) -> Value:
    if len(args) != 2:
        return VALUE
    number = _to_number(args[0])
    digits = _to_number(args[1])
    if isinstance(number, ExcelError):
        return number
    if isinstance(digits, ExcelError):
        return digits
    places = int(digits)
    factor = 10.0**places
    # Excel rounds half away from zero, which Python's round() does not do.
    scaled = abs(float(number)) * factor
    rounded = int(scaled + 0.5)
    result = rounded / factor
    return -result if float(number) < 0 else result


def _fn_if(args: Sequence[Value]) -> Value:
    if len(args) not in (2, 3):
        return VALUE
    condition = args[0]
    if isinstance(condition, ExcelError):
        return condition
    truthy = bool(_to_number(condition)) if not isinstance(condition, str) else bool(condition)
    if truthy:
        return args[1]
    return args[2] if len(args) == 3 else False


def _fn_concat(args: Sequence[Value]) -> Value:
    return "".join(_to_text(value) for value in _flatten(args))


def _fn_iserror(args: Sequence[Value]) -> Value:
    return len(args) == 1 and isinstance(args[0], ExcelError)


def _fn_abs(args: Sequence[Value]) -> Value:
    if len(args) != 1:
        return VALUE
    number = _to_number(args[0])
    return number if isinstance(number, ExcelError) else abs(float(number))


FUNCTIONS: Dict[str, Callable[[Sequence[Value]], Value]] = {
    "SUM": _fn_sum,
    "ROUND": _fn_round,
    "IF": _fn_if,
    "CONCAT": _fn_concat,
    "CONCATENATE": _fn_concat,
    "ISERROR": _fn_iserror,
    "ABS": _fn_abs,
}


def _eval(node: Node, cells: Dict[str, Value] | None = None) -> Value:
    cells = cells or {}

    if isinstance(node, Range):
        return _range_values(node, cells)

    if isinstance(node, Literal):
        return node.value

    if isinstance(node, Reference):
        return cells.get(node.name, None)

    if isinstance(node, Unary):
        value = _to_number(_eval(node.operand, cells))
        if isinstance(value, ExcelError):
            return value
        return -float(value) if node.op == "-" else float(value)

    if isinstance(node, FunctionCall):
        if node.name not in FUNCTIONS:
            return NAME
        args = [_eval(arg, cells) for arg in node.args]
        return FUNCTIONS[node.name](args)

    if isinstance(node, Binary):
        left = _eval(node.left, cells)
        right = _eval(node.right, cells)

        if node.op == "&":
            return _to_text(left) + _to_text(right)

        if node.op in _COMPARISON:
            if isinstance(left, ExcelError):
                return left
            if isinstance(right, ExcelError):
                return right
            ordering = _compare(left, right)
            return {
                "=": ordering == 0,
                "<>": ordering != 0,
                "<": ordering < 0,
                ">": ordering > 0,
                "<=": ordering <= 0,
                ">=": ordering >= 0,
            }[node.op]

        a = _to_number(left)
        b = _to_number(right)
        if isinstance(a, ExcelError):
            return a
        if isinstance(b, ExcelError):
            return b
        if node.op == "+":
            return float(a) + float(b)
        if node.op == "-":
            return float(a) - float(b)
        if node.op == "*":
            return float(a) * float(b)
        if node.op == "/":
            return DIV0 if float(b) == 0 else float(a) / float(b)
        if node.op == "^":
            return float(float(a) ** float(b))

    raise FormulaError("unreachable: unknown node")


# ---------------------------------------------------------------------------
# Ranges
# ---------------------------------------------------------------------------

_REF = re.compile("^\\$?([A-Za-z]+)\\$?(\\d+)$")


def _split_ref(ref: str) -> Tuple[str, int]:
    match = _REF.match(ref.strip())
    if match is None:
        raise FormulaError("not a cell reference: " + repr(ref))
    return match.group(1).upper(), int(match.group(2))


def _column_index(letters: str) -> int:
    """A is 1, Z is 26, AA is 27, exactly as Excel numbers columns."""
    index = 0
    for letter in letters:
        index = index * 26 + (ord(letter) - ord("A") + 1)
    return index


def _column_letters(index: int) -> str:
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def _range_values(node: "Range", cells: Dict[str, Value]) -> "RangeValues":
    start_column, start_row = _split_ref(node.start)
    end_column, end_row = _split_ref(node.end)
    first, last = _column_index(start_column), _column_index(end_column)
    if first > last:
        first, last = last, first
    top, bottom = sorted((start_row, end_row))
    values = RangeValues()
    for row in range(top, bottom + 1):
        for column in range(first, last + 1):
            values.append(cells.get(_column_letters(column) + str(row)))
    values.rows = bottom - top + 1
    values.columns = last - first + 1
    return values


# ---------------------------------------------------------------------------
# Aggregate and logical semantics
# ---------------------------------------------------------------------------


def _from_range(item: object) -> bool:
    return isinstance(item, RangeValues)


def _numeric_operands(args: Sequence[Value]) -> List[float] | ExcelError:
    """Numbers for an aggregate.

    A value read from a range contributes only when it is a number: Excel ignores text, blanks and logicals
    inside a block. A scalar argument is coerced, so SUM("abc") is #VALUE! rather than a silent zero.
    """
    numbers: List[float] = []
    for arg in args:
        if _from_range(arg):
            for item in arg:  # type: ignore[attr-defined]
                if isinstance(item, bool) or item is None:
                    continue
                if isinstance(item, (int, float)):
                    numbers.append(float(item))
            continue
        if isinstance(arg, list):
            for item in arg:
                if isinstance(item, bool) or item is None:
                    continue
                if isinstance(item, (int, float)):
                    numbers.append(float(item))
            continue
        converted = _to_number(arg)
        if isinstance(converted, ExcelError):
            return converted
        numbers.append(float(converted))
    return numbers


def _fn_min(args: Sequence[Value]) -> Value:
    numbers = _numeric_operands(_flatten(args) if not args else args)
    if isinstance(numbers, ExcelError):
        return numbers
    return min(numbers) if numbers else 0.0


def _fn_max(args: Sequence[Value]) -> Value:
    numbers = _numeric_operands(_flatten(args) if not args else args)
    if isinstance(numbers, ExcelError):
        return numbers
    return max(numbers) if numbers else 0.0


def _fn_average(args: Sequence[Value]) -> Value:
    numbers = _numeric_operands(_flatten(args) if not args else args)
    if isinstance(numbers, ExcelError):
        return numbers
    if not numbers:
        return DIV0
    return sum(numbers) / len(numbers)


def _fn_count(args: Sequence[Value]) -> Value:
    count = 0
    for arg in args:
        if isinstance(arg, list):
            for item in arg:
                if isinstance(item, (int, float)) and not isinstance(item, bool):
                    count += 1
            continue
        converted = _to_number(arg)
        if not isinstance(converted, ExcelError):
            count += 1
    return float(count)


def _fn_counta(args: Sequence[Value]) -> Value:
    count = 0
    for arg in args:
        if isinstance(arg, list):
            for item in arg:
                if item is not None and item != "":
                    count += 1
            continue
        if arg is not None and arg != "":
            count += 1
    return float(count)


def _to_bool(value: Value) -> Value:
    if isinstance(value, ExcelError):
        return value
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    text = str(value).strip().upper()
    if text == "TRUE":
        return True
    if text == "FALSE":
        return False
    return VALUE


def _logical(args: Sequence[Value], mode: str) -> Value:
    flags: List[bool] = []
    for arg in _flatten(args):
        if _from_range(arg):
            continue
        converted = _to_bool(arg)
        if isinstance(converted, ExcelError):
            return converted
        flags.append(bool(converted))
    if not flags:
        return VALUE
    if mode == "and":
        return all(flags)
    if mode == "or":
        return any(flags)
    return not flags[0]


def _fn_and(args: Sequence[Value]) -> Value:
    return _logical(args, "and")


def _fn_or(args: Sequence[Value]) -> Value:
    return _logical(args, "or")


def _fn_not(args: Sequence[Value]) -> Value:
    return _logical(args, "not")


def _fn_iferror(args: Sequence[Value]) -> Value:
    if len(args) != 2:
        return VALUE
    return args[1] if isinstance(args[0], ExcelError) else args[0]


def _fn_isblank(args: Sequence[Value]) -> Value:
    if len(args) != 1:
        return VALUE
    return args[0] is None


def _fn_isnumber(args: Sequence[Value]) -> Value:
    if len(args) != 1:
        return VALUE
    value = args[0]
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _fn_istext(args: Sequence[Value]) -> Value:
    if len(args) != 1:
        return VALUE
    return isinstance(args[0], str)


def _text_arg(args: Sequence[Value], index: int) -> str | ExcelError:
    if index >= len(args):
        return VALUE
    if isinstance(args[index], ExcelError):
        return args[index]
    return _to_text(args[index])


def _fn_len(args: Sequence[Value]) -> Value:
    text = _text_arg(args, 0)
    return text if isinstance(text, ExcelError) else float(len(text))


def _fn_left(args: Sequence[Value]) -> Value:
    text = _text_arg(args, 0)
    if isinstance(text, ExcelError):
        return text
    count = _to_number(args[1]) if len(args) > 1 else 1.0
    if isinstance(count, ExcelError):
        return count
    if count < 0:
        return VALUE
    return text[: int(count)]


def _fn_right(args: Sequence[Value]) -> Value:
    text = _text_arg(args, 0)
    if isinstance(text, ExcelError):
        return text
    count = _to_number(args[1]) if len(args) > 1 else 1.0
    if isinstance(count, ExcelError):
        return count
    if count < 0:
        return VALUE
    if int(count) == 0:
        return ""
    return text[-int(count):]


def _fn_mid(args: Sequence[Value]) -> Value:
    text = _text_arg(args, 0)
    if isinstance(text, ExcelError):
        return text
    if len(args) < 3:
        return VALUE
    start = _to_number(args[1])
    count = _to_number(args[2])
    if isinstance(start, ExcelError):
        return start
    if isinstance(count, ExcelError):
        return count
    if start < 1 or count < 0:
        return VALUE
    return text[int(start) - 1: int(start) - 1 + int(count)]


def _fn_trim(args: Sequence[Value]) -> Value:
    text = _text_arg(args, 0)
    return text if isinstance(text, ExcelError) else " ".join(text.split())


def _fn_upper(args: Sequence[Value]) -> Value:
    text = _text_arg(args, 0)
    return text if isinstance(text, ExcelError) else text.upper()


def _fn_lower(args: Sequence[Value]) -> Value:
    text = _text_arg(args, 0)
    return text if isinstance(text, ExcelError) else text.lower()


def _round_to(number: float, digits: int, mode: str) -> float:
    factor = 10 ** digits
    scaled = number * factor
    if mode == "up":
        rounded = -(-abs(scaled) // 1) * (1 if scaled >= 0 else -1)
    elif mode == "down":
        rounded = (abs(scaled) // 1) * (1 if scaled >= 0 else -1)
    else:
        return round(scaled) / factor
    return rounded / factor


def _fn_roundup(args: Sequence[Value]) -> Value:
    if len(args) != 2:
        return VALUE
    number, digits = _to_number(args[0]), _to_number(args[1])
    if isinstance(number, ExcelError):
        return number
    if isinstance(digits, ExcelError):
        return digits
    return _round_to(float(number), int(digits), "up")


def _fn_rounddown(args: Sequence[Value]) -> Value:
    if len(args) != 2:
        return VALUE
    number, digits = _to_number(args[0]), _to_number(args[1])
    if isinstance(number, ExcelError):
        return number
    if isinstance(digits, ExcelError):
        return digits
    return _round_to(float(number), int(digits), "down")


def _fn_int(args: Sequence[Value]) -> Value:
    if len(args) != 1:
        return VALUE
    number = _to_number(args[0])
    if isinstance(number, ExcelError):
        return number
    return float(int(float(number) // 1))


def _fn_mod(args: Sequence[Value]) -> Value:
    if len(args) != 2:
        return VALUE
    number, divisor = _to_number(args[0]), _to_number(args[1])
    if isinstance(number, ExcelError):
        return number
    if isinstance(divisor, ExcelError):
        return divisor
    if float(divisor) == 0:
        return DIV0
    # Excel's MOD takes the sign of the divisor, which means flooring the quotient rather than
    # truncating it: MOD(-3,2) is 1 because floor(-1.5) is -2, not -1.
    return float(number) - float(divisor) * math.floor(float(number) / float(divisor))


def _fn_sqrt(args: Sequence[Value]) -> Value:
    if len(args) != 1:
        return VALUE
    number = _to_number(args[0])
    if isinstance(number, ExcelError):
        return number
    if float(number) < 0:
        return NUM
    return float(number) ** 0.5


def _fn_sum_ranges(args: Sequence[Value]) -> Value:
    numbers = _numeric_operands(list(args))
    if isinstance(numbers, ExcelError):
        return numbers
    return float(sum(numbers))


FUNCTIONS.update(
    {
        "SUM": _fn_sum_ranges,
        "MIN": _fn_min,
        "MAX": _fn_max,
        "AVERAGE": _fn_average,
        "COUNT": _fn_count,
        "COUNTA": _fn_counta,
        "AND": _fn_and,
        "OR": _fn_or,
        "NOT": _fn_not,
        "IFERROR": _fn_iferror,
        "ISBLANK": _fn_isblank,
        "ISNUMBER": _fn_isnumber,
        "ISTEXT": _fn_istext,
        "LEN": _fn_len,
        "LEFT": _fn_left,
        "RIGHT": _fn_right,
        "MID": _fn_mid,
        "TRIM": _fn_trim,
        "UPPER": _fn_upper,
        "LOWER": _fn_lower,
        "ROUNDUP": _fn_roundup,
        "ROUNDDOWN": _fn_rounddown,
        "INT": _fn_int,
        "MOD": _fn_mod,
        "SQRT": _fn_sqrt,
    }
)


def evaluate(node: Node, cells: Dict[str, Value] | None = None) -> Value:
    """Evaluate a parsed formula in a cell context.

    A range is only meaningful to a function that consumes blocks. Evaluated on its own it is not a
    value, so it becomes #VALUE! rather than leaking a Python list into whatever called us.
    """
    result = _eval(node, cells)
    return VALUE if isinstance(result, list) else result


# ---------------------------------------------------------------------------
# Lookups: the functions reimplementations diverge on most
# ---------------------------------------------------------------------------

_NA = ExcelError("#N/A")


def _error_in(values: Sequence[Value]) -> ExcelError | None:
    for value in values:
        if isinstance(value, ExcelError):
            return value
    return None


def _as_int(value: Value) -> Value:
    """Excel truncates a fractional index rather than rounding it, and refuses a non-number."""
    number = _to_number(value)
    if isinstance(number, ExcelError):
        return number
    return int(float(number))


def _match_position(lookup: Value, values: RangeValues, match_type: int) -> Value:
    """MATCH's core: exact (0), largest <= (1, ascending), smallest >= (-1, descending)."""
    if not isinstance(values, RangeValues):
        return VALUE

    if match_type == 0:
        for index, candidate in enumerate(values, start=1):
            if _compare(candidate, lookup) == 0:
                return index
        return _NA

    best = None
    for index, candidate in enumerate(values, start=1):
        comparison = _compare(candidate, lookup)
        if match_type > 0:
            if comparison <= 0:
                best = index
            elif best is not None:
                break  # ascending: the run has passed the lookup value
        else:
            if comparison >= 0:
                best = index
            elif best is not None:
                break  # descending
    return best if best is not None else _NA


def _fn_match(args: Sequence[Value]) -> Value:
    if len(args) not in (2, 3):
        return VALUE
    error = _error_in(args)
    if error is not None:
        return error
    match_type = 1 if len(args) == 2 else _as_int(args[2])
    if isinstance(match_type, ExcelError):
        return match_type
    if match_type not in (-1, 0, 1):
        return _NA  # Excel's own behaviour for an invalid match type
    return _match_position(args[0], args[1], match_type)


def _fn_vlookup(args: Sequence[Value]) -> Value:
    """VLOOKUP(lookup, table, column, [approximate = TRUE]).

    Approximate is the *default* in Excel, and that is the trap: a table that is not sorted ascending silently
    returns a wrong row rather than an error, so the caller's data is assumed sorted, as Excel assumes it.
    """
    if len(args) not in (3, 4):
        return VALUE
    error = _error_in(args[:1])
    if error is not None:
        return error

    table = args[1]
    if not isinstance(table, RangeValues):
        return VALUE

    column = _as_int(args[2])
    if isinstance(column, ExcelError):
        return column
    if column < 1:
        return VALUE
    if column > table.columns:
        return REF

    approximate = True
    if len(args) == 4:
        flag = args[3]
        if isinstance(flag, ExcelError):
            return flag
        approximate = bool(flag)

    first_column = RangeValues(table.cell(row, 1) for row in range(1, table.rows + 1))
    first_column.rows = table.rows
    first_column.columns = 1

    position = _match_position(args[0], first_column, 1 if approximate else 0)
    if isinstance(position, ExcelError):
        return position
    return table.cell(position, column)


def _fn_index(args: Sequence[Value]) -> Value:
    """INDEX(range, row, [column]) for a 2D block, INDEX(range, position) for a single row or column.

    An array result (row 0 or column 0, the whole axis) is refused rather than approximated: the engine has no
    array values, and returning the first cell would be a confident lie.
    """
    if len(args) not in (2, 3):
        return VALUE
    block = args[0]
    if not isinstance(block, RangeValues):
        return VALUE

    is_vector = block.rows == 1 or block.columns == 1
    row = _as_int(args[1])
    if isinstance(row, ExcelError):
        return row
    if row < 1:
        return VALUE  # a 0 index means "the whole axis", which needs array values

    if is_vector and len(args) == 2:
        if row > max(block.rows, block.columns):
            return REF
        return block[row - 1]

    if len(args) != 3:
        return VALUE
    column = _as_int(args[2])
    if isinstance(column, ExcelError):
        return column
    if column < 1:
        return VALUE
    if row > block.rows or column > block.columns:
        return REF
    return block.cell(row, column)


FUNCTIONS.update({"MATCH": _fn_match, "VLOOKUP": _fn_vlookup, "INDEX": _fn_index})
