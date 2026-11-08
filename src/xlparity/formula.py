"""A spreadsheet formula parser and evaluator, starting from the parts Excel is strict about.

The reason this project exists is parity, not a roughly-right calculator. Parity starts with the details
that "Excel-ish" engines get wrong: operator precedence that matches Excel exactly, comparison rules,
error values propagating as values instead of raising, and the difference between a blank cell and an
explicit zero.

Only a handful of functions are implemented so far. The parser is the foundation; the scoreboard
grows function by function, each measured against Excel's own output rather than against a guess.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Dict, List, Sequence, Union

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

        if char in "+-*/^&%,()<>=.":
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
            return Reference(str(value))
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


def _compare(left: Value, right: Value) -> int:
    left_number = _to_number(left)
    right_number = _to_number(right)
    if not isinstance(left_number, ExcelError) and not isinstance(right_number, ExcelError):
        return (left_number > right_number) - (left_number < right_number)
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


def evaluate(node: Node, cells: Dict[str, Value] | None = None) -> Value:
    cells = cells or {}

    if isinstance(node, Literal):
        return node.value

    if isinstance(node, Reference):
        return cells.get(node.name, None)

    if isinstance(node, Unary):
        value = _to_number(evaluate(node.operand, cells))
        if isinstance(value, ExcelError):
            return value
        return -float(value) if node.op == "-" else float(value)

    if isinstance(node, FunctionCall):
        if node.name not in FUNCTIONS:
            return NAME
        args = [evaluate(arg, cells) for arg in node.args]
        return FUNCTIONS[node.name](args)

    if isinstance(node, Binary):
        left = evaluate(node.left, cells)
        right = evaluate(node.right, cells)

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
