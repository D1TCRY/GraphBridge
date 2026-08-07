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
    """Escape a string for use inside an OData single-quoted literal."""

    if not isinstance(value, str):
        raise TypeError("OData string values must be strings")
    return value.replace("'", "''")


def odata_literal(value: object) -> str:
    """Serialize a Python scalar as an OData v4 literal."""

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
    """Validate a simple OData property path and return it unchanged."""

    if not isinstance(value, str) or not value:
        raise ValueError("OData property paths cannot be empty")
    if not all(_IDENTIFIER.fullmatch(part) for part in value.split("/")):
        raise ValueError(f"invalid OData property path: {value!r}")
    return value


@dataclass(frozen=True, slots=True)
class FilterExpression:
    """A server expression with an optional equivalent local predicate."""

    expression: str
    _predicate: Callable[[Mapping[str, Any]], bool] | None = field(
        default=None, repr=False, compare=False
    )

    def to_odata(self) -> str:
        return self.expression

    def matches(self, item: Mapping[str, Any]) -> bool:
        if self._predicate is None:
            raise ValueError("this raw filter cannot be evaluated locally")
        return self._predicate(item)

    def __and__(self, other: FilterExpression) -> FilterExpression:
        predicate: Callable[[Mapping[str, Any]], bool] | None = None
        if self._predicate is not None and other._predicate is not None:
            def predicate(item: Mapping[str, Any]) -> bool:
                return self.matches(item) and other.matches(item)

        return FilterExpression(f"({self.expression}) and ({other.expression})", predicate)

    def __or__(self, other: FilterExpression) -> FilterExpression:
        predicate: Callable[[Mapping[str, Any]], bool] | None = None
        if self._predicate is not None and other._predicate is not None:
            def predicate(item: Mapping[str, Any]) -> bool:
                return self.matches(item) or other.matches(item)

        return FilterExpression(f"({self.expression}) or ({other.expression})", predicate)


def compare(field_name: str, operator: str, value: object) -> FilterExpression:
    """Build a validated comparison such as ``fields/Status eq 'Open'``."""

    path = validate_odata_path(field_name)
    normalized_operator = operator.lower()
    if normalized_operator not in _OPERATORS:
        raise ValueError(f"unsupported OData comparison operator: {operator!r}")

    def predicate(item: Mapping[str, Any]) -> bool:
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
    return compare(field_name, "eq", value)


def ne(field_name: str, value: object) -> FilterExpression:
    return compare(field_name, "ne", value)


def startswith(field_name: str, value: str) -> FilterExpression:
    """Build the ``startswith`` function supported by the list-item endpoint."""

    path = validate_odata_path(field_name)
    literal = odata_literal(value)

    def predicate(item: Mapping[str, Any]) -> bool:
        actual = _path_value(item, path)
        return isinstance(actual, str) and actual.startswith(value)

    return FilterExpression(f"startswith({path},{literal})", predicate)


def filter_from_mapping(
    values: Mapping[str, object], *, field_prefix: str | None = None
) -> FilterExpression | None:
    """Build an AND of equality comparisons from a mapping."""

    expression: FilterExpression | None = None
    for name, value in values.items():
        path = f"{field_prefix}/{name}" if field_prefix and "/" not in name else name
        part = eq(path, value)
        expression = part if expression is None else expression & part
    return expression


def fields_expand(fields: Sequence[str] | None = None) -> str:
    """Build ``fields`` or ``fields($select=...)`` for a list-item query."""

    if fields is None:
        return "fields"
    selected = [validate_odata_path(name) for name in fields]
    return f"fields($select={','.join(selected)})" if selected else "fields"


@dataclass(frozen=True, slots=True)
class ODataQuery:
    """Serialize the common stable OData query options used by GraphBridge."""

    select: tuple[str, ...] = ()
    expand: tuple[str, ...] = ()
    filter: str | FilterExpression | None = None
    top: int | None = None
    orderby: tuple[str, ...] = ()

    def to_params(self) -> dict[str, str | int]:
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
    parts = value.rsplit(" ", 1)
    path = validate_odata_path(parts[0])
    if len(parts) == 1:
        return path
    direction = parts[1].lower()
    if direction not in {"asc", "desc"}:
        raise ValueError("OData order direction must be 'asc' or 'desc'")
    return f"{path} {direction}"


def _path_value(item: Mapping[str, Any], path: str) -> Any:
    current: Any = item
    for part in path.split("/"):
        if not isinstance(current, Mapping) or part not in current:
            return None
        current = current[part]
    return current
