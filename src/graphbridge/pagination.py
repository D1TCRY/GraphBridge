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
    """Lazily retrieve and parse Graph collection pages.

    Query parameters are applied only to the first request. Each continuation
    link returned by Graph is then forwarded verbatim through the safe transport.
    No network request occurs until iteration starts.

    Args:
        transport: Transport used for collection requests.
        url: Initial collection URL.
        params: Optional parameters for the first page.
        parser: Callable that converts each raw item.

    Yields:
        Parsed collection pages.

    Raises:
        GraphInvalidResponseError: If Graph returns an invalid collection shape.
    """
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
    """Lazily yield all items from a paginated collection.

    This is a convenience layer over :func:`iter_pages` and does not materialize
    the complete collection in memory. Stopping early prevents subsequent pages
    from being requested.

    Args:
        transport: Transport used for collection requests.
        url: Initial collection URL.
        params: Optional parameters for the first page.
        parser: Callable that converts each raw item.
    """
    for page in iter_pages(transport, url, params=params, parser=parser):
        yield from page.items


def _optional_link(value: object) -> str | None:
    """Validate an optional Graph continuation link.

    A missing link ends pagination; a present link must be a non-empty string so
    the transport can validate and follow it safely.

    Args:
        value: Link value returned by Graph.

    Raises:
        GraphInvalidResponseError: If a present link is not a non-empty string.
    """
    if value is None:
        return None
    if not isinstance(value, str) or not value:
        raise GraphInvalidResponseError("Microsoft Graph returned an invalid pagination link")
    return value
