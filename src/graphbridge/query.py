"""Controlled OData helpers for stable Microsoft Graph v1.0 endpoints."""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from typing import Any

_IDENTIFIER = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_OPERATORS = frozenset({"eq", "ne", "lt", "gt", "le", "ge"})


def escape_odata_string(value: str) -> str:
    """Escape a string for an OData literal.

    Embedded single quotes are doubled according to OData rules so caller values
    cannot terminate the surrounding string literal. Prefer the controlled query
    builders over interpolating raw user input into OData text.

    Args:
        value: String to escape.

    Raises:
        TypeError: If the value is not a string.
    """

    if not isinstance(value, str):
        raise TypeError("OData string values must be strings")
    return value.replace("'", "''")


def odata_literal(value: object) -> str:
    """Serialize a Python scalar as an OData v4 literal.

    Only the scalar types deliberately supported by GraphBridge are accepted;
    arbitrary objects are never converted with an unsafe generic string cast.
    Each accepted Python type receives an explicit OData representation.

    Args:
        value: Supported scalar value to serialize.

    Raises:
        TypeError: If the value type is unsupported.
        ValueError: If a numeric value is not finite.
    """

    if value is None:
        return "null"
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, str):
        return f"'{escape_odata_string(value)}'"
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError("OData numeric values must be finite")
        return repr(value)
    raise TypeError(f"unsupported OData literal type: {type(value).__name__}")


def validate_odata_path(value: str) -> str:
    """Validate a simple OData property path.

    Every slash-separated segment must be a plain identifier, which keeps query
    builders from accepting arbitrary OData syntax in property positions. The
    original path is returned unchanged after validation.

    Args:
        value: Property path to validate.

    Raises:
        ValueError: If the path is empty or contains unsafe identifiers.
    """

    if not isinstance(value, str) or not value:
        raise ValueError("OData property paths cannot be empty")
    if not all(_IDENTIFIER.fullmatch(part) for part in value.split("/")):
        raise ValueError(f"invalid OData property path: {value!r}")
    return value


@dataclass(frozen=True, slots=True)
class FilterExpression:
    """Store an OData filter and optional local predicate.

    Controlled expressions can be sent to Graph or evaluated against downloaded
    JSON-like items. Raw filters have no local evaluator. Controlled expressions
    can be combined with ``&`` and ``|`` without losing their local predicate.

    Args:
        expression: Server-side OData expression.
        _predicate: Optional equivalent local evaluator.
    """

    expression: str
    _predicate: Callable[[Mapping[str, Any]], bool] | None = field(
        default=None, repr=False, compare=False
    )

    def to_odata(self) -> str:
        """Return the server-side OData expression.

        The value is already escaped and validated by the builder that created the
        expression. No additional escaping should be applied by the caller.
        """
        return self.expression

    def matches(self, item: Mapping[str, Any]) -> bool:
        """Evaluate the filter against one local item.

        Local evaluation is available only for expressions created by controlled
        builders and their logical combinations. Server-only raw filters raise an
        explicit error instead of being interpreted approximately.

        Args:
            item: JSON-like item to test.

        Raises:
            ValueError: If the expression has no local predicate.
        """
        if self._predicate is None:
            raise ValueError("this raw filter cannot be evaluated locally")
        return self._predicate(item)

    def __and__(self, other: FilterExpression) -> FilterExpression:
        """Combine this filter with another using logical AND.

        Both server expressions are parenthesized; local evaluation is retained
        only when both operands provide a safe predicate.

        Args:
            other: Filter to combine with this expression.
        """
        predicate: Callable[[Mapping[str, Any]], bool] | None = None
        if self._predicate is not None and other._predicate is not None:
            def predicate(item: Mapping[str, Any]) -> bool:
                """Evaluate both local predicates.

                Short-circuit evaluation mirrors the server-side logical AND.

                Args:
                    item: JSON-like item to test.
                """
                return self.matches(item) and other.matches(item)

        return FilterExpression(f"({self.expression}) and ({other.expression})", predicate)

    def __or__(self, other: FilterExpression) -> FilterExpression:
        """Combine this filter with another using logical OR.

        Both server expressions are parenthesized; local evaluation is retained
        only when both operands provide a safe predicate.

        Args:
            other: Filter to combine with this expression.
        """
        predicate: Callable[[Mapping[str, Any]], bool] | None = None
        if self._predicate is not None and other._predicate is not None:
            def predicate(item: Mapping[str, Any]) -> bool:
                """Evaluate either local predicate.

                Short-circuit evaluation mirrors the server-side logical OR.

                Args:
                    item: JSON-like item to test.
                """
                return self.matches(item) or other.matches(item)

        return FilterExpression(f"({self.expression}) or ({other.expression})", predicate)


def compare(field_name: str, operator: str, value: object) -> FilterExpression:
    """Build a validated OData comparison.

    The returned expression includes both safe server syntax and an equivalent
    local predicate for explicit local-filter mode. Operators are limited to the
    stable comparison set implemented by GraphBridge.

    Args:
        field_name: OData property path.
        operator: Supported comparison operator.
        value: Scalar value to compare.

    Raises:
        ValueError: If the field path or operator is invalid.
        TypeError: If the value cannot be represented as an OData literal.
    """

    path = validate_odata_path(field_name)
    normalized_operator = operator.lower()
    if normalized_operator not in _OPERATORS:
        raise ValueError(f"unsupported OData comparison operator: {operator!r}")

    def predicate(item: Mapping[str, Any]) -> bool:
        """Evaluate the comparison locally.

        The operator and reference value are captured from the validated builder.

        Args:
            item: JSON-like item to test.
        """
        actual = _path_value(item, path)
        if normalized_operator == "eq":
            return actual == value
        if normalized_operator == "ne":
            return actual != value
        if normalized_operator == "lt":
            return bool(actual is not None and actual < value)
        if normalized_operator == "gt":
            return bool(actual is not None and actual > value)
        if normalized_operator == "le":
            return bool(actual is not None and actual <= value)
        return bool(actual is not None and actual >= value)

    return FilterExpression(f"{path} {normalized_operator} {odata_literal(value)}", predicate)


def eq(field_name: str, value: object) -> FilterExpression:
    """Build a validated equality filter for server or local evaluation.

    This convenience wrapper delegates validation, literal formatting, and local
    predicate creation to :func:`compare`.

    Args:
        field_name: OData property path.
        value: Value that must match.
    """
    return compare(field_name, "eq", value)


def ne(field_name: str, value: object) -> FilterExpression:
    """Build a validated inequality filter for server or local evaluation.

    This convenience wrapper delegates validation, literal formatting, and local
    predicate creation to :func:`compare`.

    Args:
        field_name: OData property path.
        value: Value that must differ.
    """
    return compare(field_name, "ne", value)


def startswith(field_name: str, value: str) -> FilterExpression:
    """Build an OData ``startswith`` filter.

    The controlled expression also contains an equivalent local string-prefix
    predicate for explicit fallback mode. Non-string or missing remote values do
    not match the local predicate.

    Args:
        field_name: OData property path.
        value: Required string prefix.
    """

    path = validate_odata_path(field_name)
    literal = odata_literal(value)

    def predicate(item: Mapping[str, Any]) -> bool:
        """Evaluate the prefix test locally.

        Non-string and missing values do not satisfy the controlled expression.

        Args:
            item: JSON-like item to test.
        """
        actual = _path_value(item, path)
        return isinstance(actual, str) and actual.startswith(value)

    return FilterExpression(f"startswith({path},{literal})", predicate)


def filter_from_mapping(
    values: Mapping[str, object], *, field_prefix: str | None = None
) -> FilterExpression | None:
    """Build an AND filter from mapping entries.

    Each entry becomes an equality comparison, preserving mapping iteration order
    and optionally prefixing simple names such as list-item fields. This is the
    convenient injection-safe mapping form used by item resources.

    Args:
        values: Field names mapped to required values.
        field_prefix: Optional prefix added to simple field names.
    """

    expression: FilterExpression | None = None
    for name, value in values.items():
        path = f"{field_prefix}/{name}" if field_prefix and "/" not in name else name
        part = eq(path, value)
        expression = part if expression is None else expression & part
    return expression


def fields_expand(fields: Sequence[str] | None = None) -> str:
    """Build a list-item fields expansion.

    With selected names the result uses ``fields($select=...)``; otherwise it
    requests the complete fields relationship. Every selected name is validated
    before it becomes nested OData syntax.

    Args:
        fields: Optional field names to select.
    """

    if fields is None:
        return "fields"
    selected = [validate_odata_path(name) for name in fields]
    return f"fields($select={','.join(selected)})" if selected else "fields"


@dataclass(frozen=True, slots=True)
class ODataQuery:
    """Store supported OData query options.

    Serialization is intentionally limited to stable options used by known Graph
    endpoints, while endpoint resources decide which options they expose. The
    object itself performs no network request.

    Args:
        select: Properties included with ``$select``.
        expand: Relationships included with ``$expand``.
        filter: Optional raw or controlled filter.
        top: Optional maximum page size.
        orderby: Optional ordering expressions.
    """

    select: tuple[str, ...] = ()
    expand: tuple[str, ...] = ()
    filter: str | FilterExpression | None = None
    top: int | None = None
    orderby: tuple[str, ...] = ()

    def to_params(self) -> dict[str, str | int]:
        """Serialize configured options to request parameters.

        Empty options are omitted and property paths, page size, and ordering
        expressions are validated before any HTTP request is made. Empty options
        are omitted to keep the request minimal.

        Raises:
            ValueError: If a path, page size, or ordering is invalid.
        """
        params: dict[str, str | int] = {}
        if self.select:
            params["$select"] = ",".join(validate_odata_path(value) for value in self.select)
        if self.expand:
            params["$expand"] = ",".join(self.expand)
        if self.filter is not None:
            params["$filter"] = (
                self.filter.to_odata()
                if isinstance(self.filter, FilterExpression)
                else self.filter
            )
        if self.top is not None:
            if self.top <= 0:
                raise ValueError("top must be greater than zero")
            params["$top"] = self.top
        if self.orderby:
            params["$orderby"] = ",".join(_validate_orderby(value) for value in self.orderby)
        return params


def _validate_orderby(value: str) -> str:
    """Validate one OData ordering expression.

    Only a property path and optional ``asc`` or ``desc`` direction are accepted;
    arbitrary clauses and function calls are rejected.

    Args:
        value: Property path with an optional direction.

    Raises:
        ValueError: If the path or direction is invalid.
    """
    parts = value.rsplit(" ", 1)
    path = validate_odata_path(parts[0])
    if len(parts) == 1:
        return path
    direction = parts[1].lower()
    if direction not in {"asc", "desc"}:
        raise ValueError("OData order direction must be 'asc' or 'desc'")
    return f"{path} {direction}"


def _path_value(item: Mapping[str, Any], path: str) -> Any:
    """Read a slash-separated nested value from a JSON-like mapping.

    Missing or non-mapping intermediate values return ``None`` for predictable
    local-filter behavior.

    Args:
        item: Mapping to traverse.
        path: Slash-separated property path.
    """
    current: Any = item
    for part in path.split("/"):
        if not isinstance(current, Mapping) or part not in current:
            return None
        current = current[part]
    return current
