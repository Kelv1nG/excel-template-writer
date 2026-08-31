"""A small, safe expression language with an explicit syntax tree."""

from __future__ import annotations

import ast as python_ast
import math
import re
from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass
from decimal import Decimal
from enum import Enum, auto
from typing import Any, cast

from excel_template_writer.date_formats import (
    DateFormat,
    DateFormatValueError,
    format_date,
    parse_date_format,
)


class ExpressionSyntaxError(ValueError):
    def __init__(self, message: str, position: int = 0) -> None:
        """Create a syntax error at a character offset.

        Args:
            message: Human-readable syntax failure.
            position: Zero-based offset in the expression source.
        """

        self.position = position
        super().__init__(message)


class ExpressionEvaluationError(ValueError):
    pass


class FilterValidationError(ValueError):
    """A parsed filter name or literal argument contract is invalid."""


class FilterTypeError(ExpressionEvaluationError):
    """A runtime value has the wrong canonical type for a compiled filter."""


class ArithmeticTypeError(ExpressionEvaluationError):
    """A runtime value has the wrong canonical type for numeric arithmetic."""


class DivisionByZeroError(ExpressionEvaluationError):
    """A numeric division expression has a zero divisor."""


class NonFiniteExpressionNumberError(ExpressionEvaluationError):
    """An expression operation produced a non-finite float or decimal."""


class MissingValueError(ExpressionEvaluationError):
    def __init__(self, root: str, path: str) -> None:
        """Create a missing-value error with a stable root and display path.

        Args:
            root: Top-level context name whose lookup failed.
            path: Human-readable expression path that was unavailable.
        """

        self.root = root
        self.path = path
        super().__init__(f"missing value: {path}")


class Expression:
    pass


@dataclass(frozen=True)
class LiteralExpression(Expression):
    value: Any


@dataclass(frozen=True)
class NameExpression(Expression):
    name: str


@dataclass(frozen=True)
class AttributeExpression(Expression):
    value: Expression
    name: str


@dataclass(frozen=True)
class IndexExpression(Expression):
    value: Expression
    index: Expression


@dataclass(frozen=True)
class UnaryExpression(Expression):
    operator: str
    operand: Expression


@dataclass(frozen=True)
class BinaryExpression(Expression):
    left: Expression
    operator: str
    right: Expression


@dataclass(frozen=True)
class FilterExpression(Expression):
    value: Expression
    name: str
    arguments: tuple[Expression, ...]


@dataclass(frozen=True)
class DateFormatExpression(Expression):
    """A date filter whose literal format has been compiled semantically."""

    value: Expression
    date_format: DateFormat


class _TokenKind(Enum):
    IDENTIFIER = auto()
    STRING = auto()
    NUMBER = auto()
    DOT = auto()
    LEFT_BRACKET = auto()
    RIGHT_BRACKET = auto()
    LEFT_PAREN = auto()
    RIGHT_PAREN = auto()
    COMMA = auto()
    PIPE = auto()
    PLUS = auto()
    MINUS = auto()
    STAR = auto()
    SLASH = auto()
    EQ = auto()
    NE = auto()
    LT = auto()
    LE = auto()
    GT = auto()
    GE = auto()
    EOF = auto()


@dataclass(frozen=True)
class _Token:
    kind: _TokenKind
    value: str
    position: int


_TOKEN_PATTERN = re.compile(
    r"(?P<space>\s+)"
    r"|(?P<string>'(?:\\.|[^'\\])*'|\"(?:\\.|[^\"\\])*\")"
    r"|(?P<number>\d+(?:\.\d+)?)"
    r"|(?P<identifier>[A-Za-z_][A-Za-z0-9_]*)"
    r"|(?P<operator>==|!=|<=|>=|[.\[\](),|<>+*/-])"
)


def _tokenize(source: str) -> tuple[_Token, ...]:
    """Tokenize expression source and append an explicit end token.

    Args:
        source: Expression text without template delimiters.

    Returns:
        Immutable lexical tokens ending in ``EOF``.

    Raises:
        ExpressionSyntaxError: If the source contains an unsupported character.
    """

    tokens: list[_Token] = []
    position = 0
    while position < len(source):
        match = _TOKEN_PATTERN.match(source, position)
        if match is None:
            raise ExpressionSyntaxError(
                f"unexpected character {source[position]!r}", position=position
            )
        position = match.end()
        if match.lastgroup == "space":
            continue
        value = match.group()
        if match.lastgroup == "identifier":
            kind = _TokenKind.IDENTIFIER
        elif match.lastgroup == "string":
            kind = _TokenKind.STRING
        elif match.lastgroup == "number":
            kind = _TokenKind.NUMBER
        else:
            kind = {
                ".": _TokenKind.DOT,
                "[": _TokenKind.LEFT_BRACKET,
                "]": _TokenKind.RIGHT_BRACKET,
                "(": _TokenKind.LEFT_PAREN,
                ")": _TokenKind.RIGHT_PAREN,
                ",": _TokenKind.COMMA,
                "|": _TokenKind.PIPE,
                "+": _TokenKind.PLUS,
                "-": _TokenKind.MINUS,
                "*": _TokenKind.STAR,
                "/": _TokenKind.SLASH,
                "==": _TokenKind.EQ,
                "!=": _TokenKind.NE,
                "<": _TokenKind.LT,
                "<=": _TokenKind.LE,
                ">": _TokenKind.GT,
                ">=": _TokenKind.GE,
            }[value]
        tokens.append(_Token(kind, value, match.start()))
    tokens.append(_Token(_TokenKind.EOF, "", len(source)))
    return tuple(tokens)


class _Parser:
    def __init__(self, source: str) -> None:
        """Initialize a recursive-descent parser for one expression.

        Args:
            source: Expression text to tokenize and parse.
        """

        self._source = source
        self._tokens = _tokenize(source)
        self._index = 0

    @property
    def current(self) -> _Token:
        """Return the token at the parser cursor."""

        return self._tokens[self._index]

    def advance(self) -> _Token:
        """Consume and return the current token."""

        token = self.current
        self._index += 1
        return token

    def accept(self, kind: _TokenKind, value: str | None = None) -> _Token | None:
        """Consume the current token when its kind and optional value match.

        Args:
            kind: Required token kind.
            value: Optional exact token value.

        Returns:
            The consumed token, or ``None`` when it does not match.
        """

        token = self.current
        if token.kind is kind and (value is None or token.value == value):
            return self.advance()
        return None

    def expect(self, kind: _TokenKind, message: str) -> _Token:
        """Consume a required token or raise a positioned syntax error.

        Args:
            kind: Required token kind.
            message: Error message used when the token is absent.

        Returns:
            The consumed token.

        Raises:
            ExpressionSyntaxError: If the current token has another kind.
        """

        token = self.accept(kind)
        if token is None:
            raise ExpressionSyntaxError(message, self.current.position)
        return token

    def parse(self) -> Expression:
        """Parse one complete expression and require end-of-input."""

        expression = self.parse_pipeline()
        if self.current.kind is not _TokenKind.EOF:
            raise ExpressionSyntaxError(
                f"unexpected token {self.current.value!r}", self.current.position
            )
        return expression

    def parse_pipeline(self) -> Expression:
        """Parse a base expression followed by zero or more filters."""

        expression = self.parse_or()
        while self.accept(_TokenKind.PIPE):
            name = self.expect(_TokenKind.IDENTIFIER, "expected a filter name").value
            arguments: list[Expression] = []
            if self.accept(_TokenKind.LEFT_PAREN):
                if self.current.kind is not _TokenKind.RIGHT_PAREN:
                    arguments.append(self.parse_or())
                    while self.accept(_TokenKind.COMMA):
                        arguments.append(self.parse_or())
                self.expect(_TokenKind.RIGHT_PAREN, "expected ')' after filter arguments")
            expression = FilterExpression(expression, name, tuple(arguments))
        return expression

    def parse_or(self) -> Expression:
        """Parse left-associative boolean ``or`` expressions."""

        expression = self.parse_and()
        while self.current.kind is _TokenKind.IDENTIFIER and self.current.value == "or":
            self.advance()
            expression = BinaryExpression(expression, "or", self.parse_and())
        return expression

    def parse_and(self) -> Expression:
        """Parse left-associative boolean ``and`` expressions."""

        expression = self.parse_comparison()
        while self.current.kind is _TokenKind.IDENTIFIER and self.current.value == "and":
            self.advance()
            expression = BinaryExpression(expression, "and", self.parse_comparison())
        return expression

    def parse_comparison(self) -> Expression:
        """Parse supported equality and ordering comparisons."""

        expression = self.parse_additive()
        comparisons = {
            _TokenKind.EQ: "==",
            _TokenKind.NE: "!=",
            _TokenKind.LT: "<",
            _TokenKind.LE: "<=",
            _TokenKind.GT: ">",
            _TokenKind.GE: ">=",
        }
        while self.current.kind in comparisons:
            operator = comparisons[self.advance().kind]
            expression = BinaryExpression(expression, operator, self.parse_additive())
        return expression

    def parse_additive(self) -> Expression:
        """Parse left-associative numeric addition and subtraction."""

        expression = self.parse_multiplicative()
        operators = {
            _TokenKind.PLUS: "+",
            _TokenKind.MINUS: "-",
        }
        while self.current.kind in operators:
            operator = operators[self.advance().kind]
            expression = BinaryExpression(expression, operator, self.parse_multiplicative())
        return expression

    def parse_multiplicative(self) -> Expression:
        """Parse left-associative numeric multiplication and division."""

        expression = self.parse_unary()
        operators = {
            _TokenKind.STAR: "*",
            _TokenKind.SLASH: "/",
        }
        while self.current.kind in operators:
            operator = operators[self.advance().kind]
            expression = BinaryExpression(expression, operator, self.parse_unary())
        return expression

    def parse_unary(self) -> Expression:
        """Parse unary boolean and numeric operators or delegate to postfix parsing."""

        if self.current.kind is _TokenKind.IDENTIFIER and self.current.value == "not":
            self.advance()
            return UnaryExpression("not", self.parse_unary())
        if self.accept(_TokenKind.PLUS):
            return UnaryExpression("+", self.parse_unary())
        if self.accept(_TokenKind.MINUS):
            return UnaryExpression("-", self.parse_unary())
        return self.parse_postfix()

    def parse_postfix(self) -> Expression:
        """Parse mapping member and index access following a primary value."""

        expression = self.parse_primary()
        while True:
            if self.accept(_TokenKind.DOT):
                name = self.expect(_TokenKind.IDENTIFIER, "expected a member name after '.'")
                expression = AttributeExpression(expression, name.value)
            elif self.accept(_TokenKind.LEFT_BRACKET):
                index = self.parse_pipeline()
                self.expect(_TokenKind.RIGHT_BRACKET, "expected ']' after index")
                expression = IndexExpression(expression, index)
            else:
                return expression

    def parse_primary(self) -> Expression:
        """Parse literals, names, and parenthesized expressions."""

        token = self.current
        if self.accept(_TokenKind.LEFT_PAREN):
            expression = self.parse_pipeline()
            self.expect(_TokenKind.RIGHT_PAREN, "expected ')' after expression")
            return expression
        if token.kind is _TokenKind.STRING:
            self.advance()
            return LiteralExpression(python_ast.literal_eval(token.value))
        if token.kind is _TokenKind.NUMBER:
            self.advance()
            value: int | float = float(token.value) if "." in token.value else int(token.value)
            return LiteralExpression(value)
        if token.kind is _TokenKind.IDENTIFIER:
            self.advance()
            if token.value == "true":
                return LiteralExpression(True)
            if token.value == "false":
                return LiteralExpression(False)
            if token.value in {"none", "null"}:
                return LiteralExpression(None)
            if token.value in {"and", "or", "not"}:
                raise ExpressionSyntaxError(f"unexpected operator {token.value!r}", token.position)
            return NameExpression(token.value)
        raise ExpressionSyntaxError("expected an expression", token.position)


def parse_expression(source: str) -> Expression:
    """Parse source text into the safe expression AST.

    Args:
        source: Expression text without ``{{`` or ``}}`` delimiters.

    Returns:
        The root expression node.

    Raises:
        ExpressionSyntaxError: If the expression is empty or malformed.
    """

    if not source.strip():
        raise ExpressionSyntaxError("expression cannot be empty")
    return _Parser(source).parse()


_COLUMN_AGGREGATE_FILTERS = frozenset({"sum", "min", "max", "count"})

_FILTER_ARGUMENT_COUNTS: dict[str, frozenset[int]] = {
    "default": frozenset({1}),
    "string": frozenset({0}),
    "upper": frozenset({0}),
    "lower": frozenset({0}),
    "join": frozenset({0, 1}),
    "date": frozenset({1}),
    "sum": frozenset({0, 1}),
    "min": frozenset({0, 1}),
    "max": frozenset({0, 1}),
    "count": frozenset({0, 1}),
}


def _validate_filter_arguments(
    name: str,
    arguments: tuple[Expression, ...],
) -> None:
    """Validate one built-in filter's name, count, and literal arguments.

    Args:
        name: Parsed filter identifier.
        arguments: Recursively compiled argument expressions.

    Raises:
        FilterValidationError: If the filter contract is not satisfied.
    """

    if name not in _FILTER_ARGUMENT_COUNTS:
        raise FilterValidationError(f"unknown filter: {name}")
    allowed_counts = _FILTER_ARGUMENT_COUNTS[name]
    if len(arguments) not in allowed_counts:
        expected = (
            str(next(iter(allowed_counts)))
            if len(allowed_counts) == 1
            else " or ".join(str(count) for count in sorted(allowed_counts))
        )
        raise FilterValidationError(
            f"filter {name!r} expects {expected} argument(s); received {len(arguments)}"
        )
    if any(not isinstance(argument, LiteralExpression) for argument in arguments):
        raise FilterValidationError(f"filter {name!r} arguments must be literals")
    if name == "join" and arguments:
        separator = arguments[0]
        assert isinstance(separator, LiteralExpression)
        if not isinstance(separator.value, str):
            raise FilterValidationError("filter 'join' separator must be a string literal")
    if name in _COLUMN_AGGREGATE_FILTERS and arguments:
        column = arguments[0]
        assert isinstance(column, LiteralExpression)
        if not isinstance(column.value, str):
            raise FilterValidationError(f"filter {name!r} column must be a string literal")
        if column.value.startswith("_"):
            raise FilterValidationError(f"filter {name!r} column must not begin with '_'")


def _compile_expression_tree(expression: Expression) -> Expression:
    """Validate and lower filters within one parsed expression tree.

    Args:
        expression: Syntax-level expression produced by the parser.

    Returns:
        A semantically validated expression, including specialized filter nodes.

    Raises:
        FilterValidationError: If a filter name or argument contract is invalid.
        DateFormatSyntaxError: If a date filter format cannot be compiled.
        TypeError: If an unknown expression node reaches semantic compilation.
    """

    if isinstance(expression, (LiteralExpression, NameExpression)):
        return expression
    if isinstance(expression, AttributeExpression):
        return AttributeExpression(_compile_expression_tree(expression.value), expression.name)
    if isinstance(expression, IndexExpression):
        return IndexExpression(
            _compile_expression_tree(expression.value),
            _compile_expression_tree(expression.index),
        )
    if isinstance(expression, UnaryExpression):
        return UnaryExpression(expression.operator, _compile_expression_tree(expression.operand))
    if isinstance(expression, BinaryExpression):
        return BinaryExpression(
            _compile_expression_tree(expression.left),
            expression.operator,
            _compile_expression_tree(expression.right),
        )
    if isinstance(expression, FilterExpression):
        value = _compile_expression_tree(expression.value)
        arguments = tuple(_compile_expression_tree(item) for item in expression.arguments)
        _validate_filter_arguments(expression.name, arguments)
        if expression.name == "date":
            format_value = arguments[0]
            assert isinstance(format_value, LiteralExpression)
            if not isinstance(format_value.value, str):
                raise FilterValidationError("filter 'date' format must be a string literal")
            return DateFormatExpression(value, parse_date_format(format_value.value))
        return FilterExpression(value, expression.name, arguments)
    if isinstance(expression, DateFormatExpression):
        return DateFormatExpression(
            _compile_expression_tree(expression.value),
            expression.date_format,
        )
    raise TypeError(f"unsupported expression node: {type(expression).__name__}")


def compile_expression(source: str) -> Expression:
    """Parse, validate, and lower one safe template expression.

    Args:
        source: Expression text without ``{{`` or ``}}`` delimiters.

    Returns:
        A semantically compiled expression AST ready for evaluation.

    Raises:
        ExpressionSyntaxError: If the expression grammar is invalid.
        FilterValidationError: If a filter invocation is invalid.
        DateFormatSyntaxError: If a date format is invalid.
    """

    return _compile_expression_tree(parse_expression(source))


def _root_path(expression: Expression) -> tuple[str, str]:
    """Describe the root context name and display path of an access chain.

    Args:
        expression: Expression whose access path should be described.

    Returns:
        A ``(root, path)`` pair, or generic placeholders for computed values.
    """

    if isinstance(expression, NameExpression):
        return expression.name, expression.name
    if isinstance(expression, AttributeExpression):
        root, path = _root_path(expression.value)
        return root, f"{path}.{expression.name}"
    if isinstance(expression, IndexExpression):
        root, path = _root_path(expression.value)
        return root, f"{path}[...]"
    return "<expression>", "<expression>"


def _column_value_path(collection_path: str, index: int, column: str) -> str:
    """Describe one selected record value within a collection expression.

    Args:
        collection_path: Display path of the collection being reduced.
        index: Zero-based collection item index.
        column: Literal top-level record key selected by the filter.

    Returns:
        A readable path containing the collection index and selected key.
    """

    item_path = f"{collection_path}[{index}]"
    if re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", column):
        return f"{item_path}.{column}"
    return f"{item_path}[{column!r}]"


type _NumericValue = int | float | Decimal


class _NumericKind(Enum):
    INTEGER = auto()
    FLOAT = auto()
    DECIMAL = auto()


def _selected_collection_values(
    value: Any,
    column: str | None,
    collection_expression: Expression,
    filter_name: str,
) -> Iterator[tuple[Any, str]]:
    """Yield values selected by an aggregate filter with display paths.

    Args:
        value: Evaluated collection supplied to the aggregate filter.
        column: Optional literal top-level record key to select from every item.
        collection_expression: Input expression used to describe failing value paths.
        filter_name: Aggregate filter name used in diagnostics.

    Yields:
        Selected values paired with readable input paths.

    Raises:
        FilterTypeError: If the input is not ordered or a column item is not a record.
        MissingValueError: If a selected record does not contain ``column``.
    """

    if not isinstance(value, (list, tuple)):
        raise FilterTypeError(
            f"filter {filter_name!r} requires an ordered collection; "
            f"received {type(value).__name__}"
        )

    root, collection_path = _root_path(collection_expression)
    for index, item in enumerate(value):
        selected = item
        selected_path = f"{collection_path}[{index}]"
        if column is not None:
            selected_path = _column_value_path(collection_path, index, column)
            if not isinstance(item, Mapping):
                raise FilterTypeError(
                    f"filter {filter_name!r} with column {column!r} requires record items; "
                    f"{collection_path}[{index}] is {type(item).__name__}"
                )
            if column not in item:
                raise MissingValueError(root, selected_path)
            selected = item[column]
        yield selected, selected_path


def _numeric_kind(value: _NumericValue) -> _NumericKind:
    """Classify one already-validated canonical numeric value.

    Args:
        value: Integer, float, or decimal value to classify.

    Returns:
        The value's numeric family.
    """

    if isinstance(value, Decimal):
        return _NumericKind.DECIMAL
    if isinstance(value, float):
        return _NumericKind.FLOAT
    return _NumericKind.INTEGER


def _coerce_numeric(
    value: _NumericValue,
    kind: _NumericKind,
    description: str,
) -> _NumericValue:
    """Promote a number to a compatible aggregate or arithmetic family.

    Args:
        value: Canonical number to promote.
        kind: Destination numeric family.
        description: Operation description used for conversion failures.

    Returns:
        The promoted number.
    """

    try:
        if kind is _NumericKind.DECIMAL:
            return value if isinstance(value, Decimal) else Decimal(value)
        if kind is _NumericKind.FLOAT:
            return value if isinstance(value, float) else float(value)
        return value
    except (OverflowError, ValueError) as error:
        raise NonFiniteExpressionNumberError(
            f"{description} could not promote an operand to a finite number"
        ) from error


def _apply_promoted_numeric_operator(
    left: _NumericValue,
    operator: str,
    right: _NumericValue,
) -> _NumericValue:
    """Apply an operator after both operands have one compatible family.

    Args:
        left: Promoted left numeric operand.
        operator: Binary ``+``, ``-``, ``*``, or ``/``.
        right: Promoted right numeric operand.

    Returns:
        The numeric operation result.
    """

    left_operand = cast(Any, left)
    right_operand = cast(Any, right)
    if operator == "+":
        result = left_operand + right_operand
    elif operator == "-":
        result = left_operand - right_operand
    elif operator == "*":
        result = left_operand * right_operand
    elif operator == "/":
        result = left_operand / right_operand
    else:
        raise TypeError(f"unsupported arithmetic operator: {operator}")
    return cast(_NumericValue, result)


def _ensure_finite_expression_number(value: _NumericValue, description: str) -> _NumericValue:
    """Require a derived float or decimal to remain finite.

    Args:
        value: Derived numeric expression result.
        description: Operation description used in the diagnostic message.

    Returns:
        The unchanged finite value.

    Raises:
        NonFiniteExpressionNumberError: If ``value`` is an infinity or NaN.
    """

    if isinstance(value, float) and not math.isfinite(value):
        raise NonFiniteExpressionNumberError(f"{description} produced a non-finite float")
    if isinstance(value, Decimal) and not value.is_finite():
        raise NonFiniteExpressionNumberError(f"{description} produced a non-finite decimal")
    return value


def _require_filter_numeric(value: Any, filter_name: str, path: str) -> _NumericValue:
    """Validate one selected aggregate value as a canonical number.

    Args:
        value: Selected non-null aggregate value.
        filter_name: Aggregate filter name used in diagnostics.
        path: Readable input path for the selected value.

    Returns:
        The validated integer, float, or decimal.

    Raises:
        FilterTypeError: If the selected value is not numeric.
    """

    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        raise FilterTypeError(
            f"filter {filter_name!r} requires numeric values; {path} is {type(value).__name__}"
        )
    return value


def _merge_filter_numeric_kind(
    current: _NumericKind,
    incoming: _NumericKind,
    filter_name: str,
    path: str,
) -> _NumericKind:
    """Resolve aggregate numeric promotion or reject float/decimal mixing.

    Args:
        current: Numeric family accumulated so far.
        incoming: Family of the next selected value.
        filter_name: Aggregate filter name used in diagnostics.
        path: Readable input path of the next value.

    Returns:
        The common numeric family.

    Raises:
        FilterTypeError: If float and decimal families would be mixed.
    """

    if {current, incoming} == {_NumericKind.FLOAT, _NumericKind.DECIMAL}:
        raise FilterTypeError(
            f"filter {filter_name!r} cannot mix floating-point and decimal values; "
            f"conflict at {path}"
        )
    if _NumericKind.DECIMAL in {current, incoming}:
        return _NumericKind.DECIMAL
    if _NumericKind.FLOAT in {current, incoming}:
        return _NumericKind.FLOAT
    return _NumericKind.INTEGER


def _evaluate_numeric_aggregate_filter(
    filter_name: str,
    value: Any,
    column: str | None,
    collection_expression: Expression,
) -> _NumericValue | None:
    """Evaluate ``sum``, ``min``, or ``max`` over selected numeric values.

    Args:
        filter_name: One supported numeric aggregate filter name.
        value: Evaluated ordered collection supplied to the filter.
        column: Optional literal top-level record key.
        collection_expression: Input expression used for value paths.

    Returns:
        A promoted numeric aggregate, zero for an empty sum, or null for an empty extremum.

    Raises:
        FilterTypeError: If selected values violate the numeric aggregate contract.
        MissingValueError: If a selected record column is absent.
        NonFiniteExpressionNumberError: If summation produces a non-finite result.
    """

    aggregate: _NumericValue | None = None
    kind: _NumericKind | None = None
    for selected, selected_path in _selected_collection_values(
        value,
        column,
        collection_expression,
        filter_name,
    ):
        if selected is None:
            continue
        numeric = _require_filter_numeric(selected, filter_name, selected_path)
        incoming_kind = _numeric_kind(numeric)
        if aggregate is None or kind is None:
            aggregate = numeric
            kind = incoming_kind
            continue

        merged_kind = _merge_filter_numeric_kind(
            kind,
            incoming_kind,
            filter_name,
            selected_path,
        )
        description = f"filter {filter_name!r}"
        aggregate = _coerce_numeric(aggregate, merged_kind, description)
        numeric = _coerce_numeric(numeric, merged_kind, description)
        kind = merged_kind
        try:
            if filter_name == "sum":
                aggregate = _apply_promoted_numeric_operator(aggregate, "+", numeric)
            elif filter_name == "min":
                aggregate = numeric if numeric < aggregate else aggregate
            elif filter_name == "max":
                aggregate = numeric if numeric > aggregate else aggregate
            else:
                raise TypeError(f"unsupported numeric aggregate: {filter_name}")
        except ArithmeticError as error:
            raise NonFiniteExpressionNumberError(
                f"filter {filter_name!r} could not produce a finite number"
            ) from error
        aggregate = _ensure_finite_expression_number(
            aggregate,
            f"filter {filter_name!r}",
        )

    if aggregate is None:
        return 0 if filter_name == "sum" else None
    return aggregate


def _evaluate_count_filter(
    value: Any,
    column: str | None,
    collection_expression: Expression,
) -> int:
    """Count collection items or present non-null record-column values.

    Args:
        value: Evaluated ordered collection supplied to ``count``.
        column: Optional literal top-level record key.
        collection_expression: Input expression used for value paths.

    Returns:
        Collection length without a column, otherwise the non-null selected-value count.

    Raises:
        FilterTypeError: If the input or column-mode row shape is invalid.
        MissingValueError: If a selected record column is absent.
    """

    if column is None:
        if not isinstance(value, (list, tuple)):
            raise FilterTypeError(
                f"filter 'count' requires an ordered collection; received {type(value).__name__}"
            )
        return len(value)
    return sum(
        selected is not None
        for selected, _ in _selected_collection_values(
            value,
            column,
            collection_expression,
            "count",
        )
    )


def _require_arithmetic_numeric(value: Any, operator: str, operand: str) -> _NumericValue:
    """Validate one non-null arithmetic operand as numeric.

    Args:
        value: Runtime operand value.
        operator: Arithmetic operator used in diagnostics.
        operand: Operand label such as ``left`` or ``right``.

    Returns:
        The validated integer, float, or decimal.

    Raises:
        ArithmeticTypeError: If ``value`` is not numeric.
    """

    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        raise ArithmeticTypeError(
            f"operator {operator!r} requires numeric operands; "
            f"{operand} operand is {type(value).__name__}"
        )
    return value


def _arithmetic_numeric_kind(
    left: _NumericValue,
    right: _NumericValue,
    operator: str,
) -> _NumericKind:
    """Resolve a binary arithmetic numeric family.

    Args:
        left: Validated left numeric operand.
        right: Validated right numeric operand.
        operator: Arithmetic operator used in diagnostics.

    Returns:
        The common numeric family for the operation.

    Raises:
        ArithmeticTypeError: If float and decimal operands would be mixed.
    """

    left_kind = _numeric_kind(left)
    right_kind = _numeric_kind(right)
    if {left_kind, right_kind} == {_NumericKind.FLOAT, _NumericKind.DECIMAL}:
        raise ArithmeticTypeError(
            f"operator {operator!r} cannot mix floating-point and decimal operands"
        )
    if _NumericKind.DECIMAL in {left_kind, right_kind}:
        return _NumericKind.DECIMAL
    if _NumericKind.FLOAT in {left_kind, right_kind}:
        return _NumericKind.FLOAT
    return _NumericKind.INTEGER


def _evaluate_unary_arithmetic(operator: str, value: Any) -> _NumericValue | None:
    """Evaluate a numeric unary sign with null propagation.

    Args:
        operator: Unary ``+`` or ``-``.
        value: Evaluated operand value.

    Returns:
        Signed numeric value, or null when the operand is null.

    Raises:
        ArithmeticTypeError: If the operand is not numeric.
        NonFiniteExpressionNumberError: If the result is non-finite.
    """

    if value is None:
        return None
    numeric = _require_arithmetic_numeric(value, operator, "unary")
    try:
        result = +numeric if operator == "+" else -numeric
    except ArithmeticError as error:
        raise NonFiniteExpressionNumberError(
            f"unary operator {operator!r} could not produce a finite number"
        ) from error
    return _ensure_finite_expression_number(result, f"unary operator {operator!r}")


def _evaluate_binary_arithmetic(
    left: Any,
    operator: str,
    right: Any,
) -> _NumericValue | None:
    """Evaluate one numeric binary operation with promotion and null propagation.

    Args:
        left: Evaluated left operand.
        operator: Binary ``+``, ``-``, ``*``, or ``/``.
        right: Evaluated right operand.

    Returns:
        Promoted numeric result, or null when either operand is null.

    Raises:
        ArithmeticTypeError: If an operand is not numeric or families cannot mix.
        DivisionByZeroError: If the divisor is zero.
        NonFiniteExpressionNumberError: If the result is non-finite.
    """

    if left is None or right is None:
        return None
    left_numeric = _require_arithmetic_numeric(left, operator, "left")
    right_numeric = _require_arithmetic_numeric(right, operator, "right")
    kind = _arithmetic_numeric_kind(left_numeric, right_numeric, operator)
    description = f"operator {operator!r}"
    left_numeric = _coerce_numeric(left_numeric, kind, description)
    right_numeric = _coerce_numeric(right_numeric, kind, description)
    if operator == "/" and right_numeric == 0:
        raise DivisionByZeroError("division by zero")

    try:
        result = _apply_promoted_numeric_operator(left_numeric, operator, right_numeric)
    except ArithmeticError as error:
        raise NonFiniteExpressionNumberError(
            f"operator {operator!r} could not produce a finite number"
        ) from error
    return _ensure_finite_expression_number(result, f"operator {operator!r}")


def _evaluate(expression: Expression, scope: Mapping[str, Any]) -> Any:
    """Evaluate one expression node against a canonical lexical scope.

    Args:
        expression: Expression node to evaluate.
        scope: Current canonical variable mapping.

    Returns:
        The resulting canonical or intermediate scalar value.

    Raises:
        ExpressionEvaluationError: If access or a filter is unsupported.
        MissingValueError: If a referenced value cannot be found.
        TypeError: If an unknown AST node is encountered.
    """

    if isinstance(expression, LiteralExpression):
        return expression.value
    if isinstance(expression, NameExpression):
        if expression.name.startswith("_"):
            raise ExpressionEvaluationError("access to private names is forbidden")
        if expression.name not in scope:
            raise MissingValueError(expression.name, expression.name)
        return scope[expression.name]
    if isinstance(expression, AttributeExpression):
        if expression.name.startswith("_"):
            raise ExpressionEvaluationError("access to private members is forbidden")
        value = _evaluate(expression.value, scope)
        root, path = _root_path(expression)
        if not isinstance(value, Mapping) or expression.name not in value:
            raise MissingValueError(root, path)
        return value[expression.name]
    if isinstance(expression, IndexExpression):
        value = _evaluate(expression.value, scope)
        index = _evaluate(expression.index, scope)
        root, path = _root_path(expression)
        if isinstance(index, str) and index.startswith("_"):
            raise ExpressionEvaluationError("access to private members is forbidden")
        try:
            return value[index]
        except (IndexError, KeyError, TypeError) as error:
            raise MissingValueError(root, path) from error
    if isinstance(expression, UnaryExpression):
        value = _evaluate(expression.operand, scope)
        if expression.operator == "not":
            return not bool(value)
        if expression.operator in {"+", "-"}:
            return _evaluate_unary_arithmetic(expression.operator, value)
        raise TypeError(f"unsupported unary operator: {expression.operator}")
    if isinstance(expression, BinaryExpression):
        left = _evaluate(expression.left, scope)
        if expression.operator == "and":
            return bool(left) and bool(_evaluate(expression.right, scope))
        if expression.operator == "or":
            return bool(left) or bool(_evaluate(expression.right, scope))
        right = _evaluate(expression.right, scope)
        if expression.operator in {"+", "-", "*", "/"}:
            return _evaluate_binary_arithmetic(left, expression.operator, right)
        return {
            "==": lambda: left == right,
            "!=": lambda: left != right,
            "<": lambda: left < right,
            "<=": lambda: left <= right,
            ">": lambda: left > right,
            ">=": lambda: left >= right,
        }[expression.operator]()
    if isinstance(expression, DateFormatExpression):
        value = _evaluate(expression.value, scope)
        try:
            return format_date(value, expression.date_format)
        except DateFormatValueError as error:
            raise FilterTypeError(str(error)) from error
    if isinstance(expression, FilterExpression):
        try:
            value = _evaluate(expression.value, scope)
        except MissingValueError:
            if expression.name != "default" or not expression.arguments:
                raise
            return _evaluate(expression.arguments[0], scope)
        arguments = [_evaluate(argument, scope) for argument in expression.arguments]
        if expression.name == "default":
            return value
        if expression.name == "string":
            return "" if value is None else str(value)
        if expression.name == "upper":
            return str(value).upper()
        if expression.name == "lower":
            return str(value).lower()
        if expression.name == "join":
            separator = str(arguments[0]) if arguments else ", "
            return separator.join(str(item) for item in value)
        if expression.name in {"sum", "min", "max"}:
            column = arguments[0] if arguments else None
            assert column is None or isinstance(column, str)
            return _evaluate_numeric_aggregate_filter(
                expression.name,
                value,
                column,
                expression.value,
            )
        if expression.name == "count":
            column = arguments[0] if arguments else None
            assert column is None or isinstance(column, str)
            return _evaluate_count_filter(value, column, expression.value)
        raise ExpressionEvaluationError(f"unknown filter: {expression.name}")
    raise TypeError(f"unsupported expression node: {type(expression).__name__}")


def evaluate_expression(expression: Expression, scope: Mapping[str, Any]) -> Any:
    """Evaluate without attribute access, calls, imports, or Python execution.

    Args:
        expression: Parsed expression AST.
        scope: Canonical lexical scope visible to the expression.

    Returns:
        The evaluated value.

    Raises:
        ExpressionEvaluationError: If evaluation violates the safe expression contract.
    """

    return _evaluate(expression, scope)


def expression_root_names(expression: Expression) -> Sequence[str]:
    """Return unique root names read by an expression.

    Args:
        expression: Parsed expression tree to inspect.

    Returns:
        Root names in first-seen traversal order.
    """

    names: list[str] = []
    if isinstance(expression, NameExpression):
        names.append(expression.name)
    for value in vars(expression).values():
        if isinstance(value, Expression):
            names.extend(expression_root_names(value))
        elif isinstance(value, tuple):
            for item in value:
                if isinstance(item, Expression):
                    names.extend(expression_root_names(item))
    return tuple(dict.fromkeys(names))
