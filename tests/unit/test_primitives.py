from __future__ import annotations

from typing import Any

import pytest

from graphbridge.batch import BatchRequest, batch_payload, chunks
from graphbridge.models import (
    BatchResult,
    ColumnInfo,
    GraphError,
    ListInfo,
    ListItem,
    Page,
    SiteInfo,
    SyncPlan,
    SyncResult,
)
from graphbridge.pagination import iter_items
from graphbridge.query import ODataQuery


def test_lightweight_models_preserve_typed_fields_and_raw_payloads() -> None:
    site = SiteInfo.from_payload({"id": 1, "displayName": "Site", "webUrl": "https://example.invalid"})
    sharepoint_list = ListInfo.from_payload({"id": 2, "name": "Tasks"})
    item = ListItem.from_payload(
        {"id": 3, "fields": {"Title": "One"}, "@odata.etag": "etag"}
    )
    column = ColumnInfo.from_payload(
        {"id": 4, "name": "Title", "displayName": "Title", "description": "Text"}
    )
    error = GraphError(code="invalid", message="Invalid", status_code=400)
    page = Page(items=[item], next_link="https://next.example.invalid")
    batch = BatchResult(successes=[item], failures=[error])
    plan = SyncPlan(creates=[{"Title": "New"}], updates=[], deletes=["1"])
    result = SyncResult(created=[item], failures=[error])

    assert (site.id, site.display_name, site.web_url) == ("1", "Site", "https://example.invalid")
    assert (sharepoint_list.id, sharepoint_list.name) == ("2", "Tasks")
    assert (item.id, item.fields, item.etag) == ("3", {"Title": "One"}, "etag")
    assert (column.id, column.description) == ("4", "Text")
    assert page.items == batch.successes == result.created
    assert batch.failures == result.failures == [error]
    assert plan.deletes == ["1"]


def test_odata_query_serializes_supported_v1_parameters() -> None:
    query = ODataQuery(
        select=("id", "displayName"),
        expand=("fields",),
        filter="fields/Status eq 'Open'",
        top=50,
    )

    assert query.to_params() == {
        "$select": "id,displayName",
        "$expand": "fields",
        "$filter": "fields/Status eq 'Open'",
        "$top": 50,
    }
    with pytest.raises(ValueError, match="greater than zero"):
        ODataQuery(top=0).to_params()


def test_batch_primitives_chunk_and_validate_graph_limit() -> None:
    request = BatchRequest(
        id="1",
        method="post",
        url="/sites/site/lists/list/items",
        headers={"Content-Type": "application/json"},
        body={"fields": {"Title": "One"}},
    )

    assert request.to_payload() == {
        "id": "1",
        "method": "POST",
        "url": "/sites/site/lists/list/items",
        "headers": {"Content-Type": "application/json"},
        "body": {"fields": {"Title": "One"}},
    }
    assert [list(part) for part in chunks(list(range(21)))] == [list(range(20)), [20]]
    with pytest.raises(ValueError, match="between 1 and 20"):
        list(chunks([1], size=0))
    with pytest.raises(ValueError, match="at most 20"):
        batch_payload([request] * 21)


def test_iter_items_uses_page_parser_and_exact_next_link() -> None:
    next_link = "https://next.example.invalid/items?$skiptoken=opaque%2Bvalue"

    class FakeTransport:
        def __init__(self) -> None:
            self.calls: list[tuple[str, object]] = []
            self.payloads = [
                {"value": [{"id": "1"}], "@odata.nextLink": next_link},
                {"value": [{"id": "2"}]},
            ]

        def get(self, url: str, *, params: object = None) -> Any:
            self.calls.append((url, params))
            return self.payloads.pop(0)

    transport = FakeTransport()
    items = list(
        iter_items(
            transport,  # type: ignore[arg-type]
            "/items",
            params={"$top": 1},
            parser=ListItem.from_payload,
        )
    )

    assert [item.id for item in items] == ["1", "2"]
    assert transport.calls == [("/items", {"$top": 1}), (next_link, None)]
