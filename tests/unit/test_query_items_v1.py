from __future__ import annotations

import json
from datetime import date, datetime, timezone
from decimal import Decimal
from types import SimpleNamespace
from urllib.parse import parse_qs, urlsplit

import pytest
import responses

from graphbridge import GraphBridgeClient
from graphbridge.exceptions import GraphPreconditionFailedError
from graphbridge.models import ListInfo, SiteInfo
from graphbridge.query import (
    FilterExpression,
    ODataQuery,
    compare,
    fields_expand,
    filter_from_mapping,
    odata_literal,
    startswith,
    validate_odata_path,
)

BASE_URL = "https://graph.query.invalid/v1.0"
SITE_ID = "query.invalid,collection,site"
LIST_ID = "query-list"


class Credential:
    def get_token(self, _scope: str) -> SimpleNamespace:
        return SimpleNamespace(token="query-token")


def tasks() -> object:
    client = GraphBridgeClient(credential=Credential(), base_url=BASE_URL, max_retries=0)
    site = client.sites.bind(SiteInfo(id=SITE_ID))
    return site.lists.bind(ListInfo(id=LIST_ID))


def request_json(index: int) -> object:
    body = responses.calls[index].request.body
    if isinstance(body, bytes):
        body = body.decode("utf-8")
    return json.loads(body)


def test_controlled_odata_query_escapes_selects_expands_and_orders() -> None:
    expression = filter_from_mapping({"Title": "O'Brien"}, field_prefix="fields")
    assert expression is not None
    assert expression.to_odata() == "fields/Title eq 'O''Brien'"
    assert startswith("fields/Title", "D'").to_odata() == "startswith(fields/Title,'D''')"
    assert fields_expand(("Title", "Status")) == "fields($select=Title,Status)"
    assert ODataQuery(
        select=("id", "displayName"),
        expand=("columns",),
        filter=expression,
        top=10,
        orderby=("displayName desc",),
    ).to_params() == {
        "$select": "id,displayName",
        "$expand": "columns",
        "$filter": "fields/Title eq 'O''Brien'",
        "$top": 10,
        "$orderby": "displayName desc",
    }


def test_odata_literals_validation_and_local_predicates() -> None:
    assert [
        odata_literal(None),
        odata_literal(True),
        odata_literal(False),
        odata_literal(3),
        odata_literal(Decimal("2.50")),
        odata_literal(date(2026, 1, 2)),
        odata_literal(datetime(2026, 1, 2, tzinfo=timezone.utc)),
    ] == ["null", "true", "false", "3", "2.50", "2026-01-02", "2026-01-02T00:00:00+00:00"]
    item = {"fields": {"Priority": 3, "Title": "Quarterly"}}
    assert compare("fields/Priority", "ge", 3).matches(item)
    assert compare("fields/Priority", "lt", 4).matches(item)
    assert compare("fields/Priority", "ne", 2).matches(item)
    assert (
        compare("fields/Priority", "eq", 3) & startswith("fields/Title", "Quart")
    ).matches(item)
    assert (
        compare("fields/Priority", "eq", 9) | startswith("fields/Title", "Quart")
    ).matches(item)
    with pytest.raises(ValueError, match="finite"):
        odata_literal(float("nan"))
    with pytest.raises(TypeError, match="unsupported"):
        odata_literal(object())
    with pytest.raises(ValueError, match="invalid OData"):
        validate_odata_path("fields/Title eq hacked")
    with pytest.raises(ValueError, match="unsupported"):
        compare("fields/Title", "contains", "x")
    with pytest.raises(ValueError, match="cannot be evaluated"):
        FilterExpression("raw eq true").matches({})
    with pytest.raises(ValueError, match="direction"):
        ODataQuery(orderby=("displayName sideways",)).to_params()


@responses.activate
def test_item_mapping_filter_runs_on_server_by_default() -> None:
    resource = tasks()
    items_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/items"
    responses.get(items_url, json={"value": []}, status=200)

    resource.items.list(fields=("Title",), filter={"Title": "O'Brien"}, top=5)

    query = parse_qs(urlsplit(responses.calls[0].request.url).query)
    assert query == {
        "$expand": ["fields($select=Title)"],
        "$filter": ["fields/Title eq 'O''Brien'"],
        "$top": ["5"],
    }


@responses.activate
def test_local_filter_is_explicit_and_downloads_each_page_once() -> None:
    resource = tasks()
    items_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/items"
    next_link = f"{items_url}?$skiptoken=opaque%2Bnext"
    responses.get(
        items_url,
        json={
            "value": [
                {"id": "1", "fields": {"Status": "Open"}},
                {"id": "2", "fields": {"Status": "Done"}},
            ],
            "@odata.nextLink": next_link,
        },
        status=200,
    )
    responses.get(
        next_link,
        json={"value": [{"id": "3", "fields": {"Status": "Open"}}]},
        status=200,
    )

    matched = list(
        resource.items.iter_all(
            filter={"Status": "Open"},
            filter_mode="local",
        )
    )

    assert [item.id for item in matched] == ["1", "3"]
    assert len(responses.calls) == 2
    assert "$filter" not in parse_qs(urlsplit(responses.calls[0].request.url).query)
    assert responses.calls[1].request.url == next_link
    with pytest.raises(ValueError, match="raw string filters"):
        list(resource.items.iter_all(filter="fields/Status eq 'Open'", filter_mode="local"))


@responses.activate
def test_iter_all_is_lazy_and_does_not_fetch_next_page_early() -> None:
    resource = tasks()
    items_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/items"
    next_link = f"{items_url}?$skiptoken=second"
    responses.get(
        items_url,
        json={"value": [{"id": "1", "fields": {}}], "@odata.nextLink": next_link},
        status=200,
    )
    responses.get(next_link, json={"value": [{"id": "2", "fields": {}}]}, status=200)

    iterator = resource.items.iter_all()
    assert len(responses.calls) == 0
    assert next(iterator).id == "1"
    assert len(responses.calls) == 1
    assert next(iterator).id == "2"
    assert len(responses.calls) == 2


@responses.activate
def test_item_crud_preserves_models_etags_and_empty_responses() -> None:
    resource = tasks()
    items_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/items"
    item_url = f"{items_url}/7"
    responses.get(
        item_url,
        json={"id": "7", "eTag": "etag-1", "fields": {"Title": "Before"}},
        status=200,
    )
    responses.post(items_url, body="", status=201)
    responses.patch(f"{item_url}/fields", body="", status=204)
    responses.delete(item_url, status=204)

    existing = resource.items.get("7")
    created = resource.items.create({"Title": "Created"})
    updated = resource.items.update("7", {"Title": "After"}, etag="etag-1")
    assert resource.items.delete("7", etag="etag-2") is None

    assert (existing.id, existing.etag) == ("7", "etag-1")
    assert created.response_empty is True and created.fields == {"Title": "Created"}
    assert updated.response_empty is True and updated.fields == {"Title": "After"}
    assert request_json(1) == {"fields": {"Title": "Created"}}
    assert responses.calls[2].request.headers["If-Match"] == "etag-1"
    assert responses.calls[3].request.headers["If-Match"] == "etag-2"


@responses.activate
def test_item_update_surfaces_412_precondition_failure() -> None:
    resource = tasks()
    update_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/items/9/fields"
    responses.patch(
        update_url,
        json={"error": {"code": "preconditionFailed", "message": "etag mismatch"}},
        status=412,
    )

    with pytest.raises(GraphPreconditionFailedError) as caught:
        resource.items.update("9", {"Title": "Nope"}, etag='W/"old"')

    assert caught.value.status_code == 412
    assert responses.calls[0].request.headers["If-Match"] == 'W/"old"'
