"""Pagination helpers that follow Microsoft Graph nextLink values verbatim."""

from __future__ import annotations

from collections.abc import Callable, Iterator, Mapping
from typing import Any, TypeVar

from .exceptions import GraphInvalidResponseError
from .models import Page
from .transport import GraphTransport

T = TypeVar("T")


def iter_pages(
    transport: GraphTransport,
    url: str,
    *,
    params: Mapping[str, str | int] | None = None,
    parser: Callable[[Mapping[str, Any]], T],
) -> Iterator[Page[T]]:
    current_url: str | None = url
    current_params = params
    while current_url is not None:
        payload = transport.get(current_url, params=current_params)
        if not isinstance(payload, Mapping):
            raise GraphInvalidResponseError("Microsoft Graph collection response must be a JSON object")
        raw_items = payload.get("value", [])
        if not isinstance(raw_items, list):
            raise GraphInvalidResponseError("Microsoft Graph collection 'value' must be a list")
        parsed_items = []
        for item in raw_items:
            if not isinstance(item, Mapping):
                raise GraphInvalidResponseError("Microsoft Graph collection items must be JSON objects")
            parsed_items.append(parser(item))
        next_link = _optional_link(payload.get("@odata.nextLink"))
        delta_link = _optional_link(payload.get("@odata.deltaLink"))
        yield Page(items=parsed_items, next_link=next_link, delta_link=delta_link, raw=dict(payload))
        current_url = next_link
        current_params = None


def iter_items(
    transport: GraphTransport,
    url: str,
    *,
    params: Mapping[str, str | int] | None = None,
    parser: Callable[[Mapping[str, Any]], T],
) -> Iterator[T]:
    for page in iter_pages(transport, url, params=params, parser=parser):
        yield from page.items


def _optional_link(value: object) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str) or not value:
        raise GraphInvalidResponseError("Microsoft Graph returned an invalid pagination link")
    return value
