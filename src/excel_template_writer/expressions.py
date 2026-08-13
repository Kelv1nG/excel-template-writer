"""A small, safe expression language with an explicit syntax tree."""

from __future__ import annotations

import ast as python_ast
import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from enum import Enum, auto
from typing import Any


class ExpressionSyntaxError(ValueError):
    def __init__(self, message: str, position: int = 0) -> None:
        self.position = position
        super().__init__(message)


class ExpressionEvaluationError(ValueError):
    pass


class MissingValueError(ExpressionEvaluationError):
    def __init__(self, root: str, path: str) -> None:
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
    r"|(?P<operator>==|!=|<=|>=|[.\[\](),|<>])"
)


def _tokenize(source: str) -> tuple[_Token, ...]:
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
        self._source = source
        self._tokens = _tokenize(source)
        self._index = 0

    @property
    def current(self) -> _Token:
        return self._tokens[self._index]

    def advance(self) -> _Token:
        token = self.current
        self._index += 1
        return token

    def accept(self, kind: _TokenKind, value: str | None = None) -> _Token | None:
        token = self.current
        if token.kind is kind and (value is None or token.value == value):
            return self.advance()
        return None

    def expect(self, kind: _TokenKind, message: str) -> _Token:
        token = self.accept(kind)
        if token is None:
            raise ExpressionSyntaxError(message, self.current.position)
        return token

    def parse(self) -> Expression:
        expression = self.parse_pipeline()
        if self.current.kind is not _TokenKind.EOF:
            raise ExpressionSyntaxError(
                f"unexpected token {self.current.value!r}", self.current.position
            )
        return expression

    def parse_pipeline(self) -> Expression:
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
        expression = self.parse_and()
        while self.current.kind is _TokenKind.IDENTIFIER and self.current.value == "or":
            self.advance()
            expression = BinaryExpression(expression, "or", self.parse_and())
        return expression

    def parse_and(self) -> Expression:
        expression = self.parse_comparison()
        while self.current.kind is _TokenKind.IDENTIFIER and self.current.value == "and":
            self.advance()
            expression = BinaryExpression(expression, "and", self.parse_comparison())
        return expression

    def parse_comparison(self) -> Expression:
        expression = self.parse_unary()
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
            expression = BinaryExpression(expression, operator, self.parse_unary())
        return expression

    def parse_unary(self) -> Expression:
        if self.current.kind is _TokenKind.IDENTIFIER and self.current.value == "not":
            self.advance()
            return UnaryExpression("not", self.parse_unary())
        return self.parse_postfix()

    def parse_postfix(self) -> Expression:
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
    if not source.strip():
        raise ExpressionSyntaxError("expression cannot be empty")
    return _Parser(source).parse()


def _root_path(expression: Expression) -> tuple[str, str]:
    if isinstance(expression, NameExpression):
        return expression.name, expression.name
    if isinstance(expression, AttributeExpression):
        root, path = _root_path(expression.value)
        return root, f"{path}.{expression.name}"
    if isinstance(expression, IndexExpression):
        root, path = _root_path(expression.value)
        return root, f"{path}[...]"
    return "<expression>", "<expression>"


def _evaluate(expression: Expression, scope: Mapping[str, Any]) -> Any:
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
        return not bool(_evaluate(expression.operand, scope))
    if isinstance(expression, BinaryExpression):
        left = _evaluate(expression.left, scope)
        if expression.operator == "and":
            return bool(left) and bool(_evaluate(expression.right, scope))
        if expression.operator == "or":
            return bool(left) or bool(_evaluate(expression.right, scope))
        right = _evaluate(expression.right, scope)
        return {
            "==": lambda: left == right,
            "!=": lambda: left != right,
            "<": lambda: left < right,
            "<=": lambda: left <= right,
            ">": lambda: left > right,
            ">=": lambda: left >= right,
        }[expression.operator]()
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
        raise ExpressionEvaluationError(f"unknown filter: {expression.name}")
    raise TypeError(f"unsupported expression node: {type(expression).__name__}")


def evaluate_expression(expression: Expression, scope: Mapping[str, Any]) -> Any:
    """Evaluate without attribute access, calls, imports, or arbitrary Python execution."""

    return _evaluate(expression, scope)


def expression_root_names(expression: Expression) -> Sequence[str]:
    """Return names read by an expression; useful for future static validation."""

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
