from __future__ import annotations

import json
from types import SimpleNamespace
from urllib.parse import parse_qs, urlsplit

import pytest
import requests
import responses

from graphbridge import GraphBridgeClient
from graphbridge.auth import GRAPH_SCOPE
from graphbridge.exceptions import GraphInvalidResponseError
from graphbridge.models import ListInfo, SiteInfo
from graphbridge.resources import (
    ColumnsResource,
    ListItemsResource,
    ListsResource,
    SharePointListResource,
    SiteResource,
    VersionsResource,
)
from graphbridge.transport import GraphTransport

BASE_URL = "https://graph.example.invalid/v1.0"
HOSTNAME = "tenant.example.invalid"
SITE_PATH = "/sites/Marketing Team"
SITE_ID = "tenant.example.invalid,collection-id,web-id"
LIST_ID = "list-id"
TOKEN = "composed-api-token"


class ReusableCredential:
    def __init__(self) -> None:
        self.calls: list[str] = []

    def get_token(self, scope: str) -> SimpleNamespace:
        self.calls.append(scope)
        return SimpleNamespace(token=TOKEN, expires_on=9999999999)

    def __repr__(self) -> str:
        return "ReusableCredential(secret=do-not-render)"


def bound_list(client: GraphBridgeClient) -> SharePointListResource:
    site = client.sites.bind(SiteInfo(id=SITE_ID, display_name="Marketing"))
    return site.lists.bind(ListInfo(id=LIST_ID, display_name="Tasks"))


def request_json(index: int) -> object:
    body = responses.calls[index].request.body
    if isinstance(body, bytes):
        body = body.decode("utf-8")
    return json.loads(body)


@responses.activate
def test_documented_composition_uses_one_client_transport_and_session() -> None:
    credential = ReusableCredential()
    session = requests.Session()
    client = GraphBridgeClient(
        credential=credential,
        session=session,
        base_url=BASE_URL,
        max_retries=0,
    )
    site_url = f"{BASE_URL}/sites/{HOSTNAME}:/sites/Marketing%20Team"
    list_url = f"{BASE_URL}/sites/{SITE_ID}/lists/Tasks"
    responses.get(site_url, json={"id": SITE_ID, "displayName": "Marketing"}, status=200)
    responses.get(list_url, json={"id": LIST_ID, "displayName": "Tasks"}, status=200)

    site = client.sites.get_by_path(hostname=HOSTNAME, path=SITE_PATH)
    tasks = site.lists.get("Tasks")

    assert site.info == SiteInfo(
        id=SITE_ID,
        display_name="Marketing",
        raw={"id": SITE_ID, "displayName": "Marketing"},
    )
    assert tasks.info == ListInfo(
        id=LIST_ID,
        display_name="Tasks",
        raw={"id": LIST_ID, "displayName": "Tasks"},
    )
    assert site.client is client
    assert tasks.client is client
    assert site.lists.client is client
    assert tasks.items.client is client
    assert tasks.columns.client is client
    assert tasks.versions.client is client
    assert all(
        resource.transport is client.transport
        for resource in (site, site.lists, tasks, tasks.items, tasks.columns, tasks.versions)
    )
    assert client.transport.session is session
    assert credential.calls == [GRAPH_SCOPE, GRAPH_SCOPE]
    assert [call.request.url for call in responses.calls] == [site_url, list_url]
    assert all("/v1.0/" in call.request.url and "/beta/" not in call.request.url for call in responses.calls)


@responses.activate
def test_item_pagination_uses_exact_next_link_and_models() -> None:
    client = GraphBridgeClient(credential=ReusableCredential(), base_url=BASE_URL, max_retries=0)
    tasks = bound_list(client)
    items_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/items"
    next_link = f"{items_url}?$skiptoken=opaque%2Bvalue&$top=2"
    responses.get(
        items_url,
        json={
            "value": [{"id": "1", "fields": {"Title": "First"}}],
            "@odata.nextLink": next_link,
        },
        status=200,
    )
    responses.get(
        next_link,
        json={"value": [{"id": "2", "fields": {"Title": "Second"}}]},
        status=200,
    )

    pages = list(
        tasks.items.iter_pages(
            fields=("Title", "Status"),
            filter="fields/Status eq 'Open'",
            top=2,
        )
    )

    assert [[item.id for item in page.items] for page in pages] == [["1"], ["2"]]
    assert pages[0].items[0].fields == {"Title": "First"}
    assert pages[0].next_link == next_link
    first_query = parse_qs(urlsplit(responses.calls[0].request.url).query)
    assert first_query == {
        "$expand": ["fields($select=Title,Status)"],
        "$filter": ["fields/Status eq 'Open'"],
        "$top": ["2"],
    }
    assert responses.calls[1].request.url == next_link


@responses.activate
def test_columns_and_item_crud_return_models_and_preserve_request_shapes() -> None:
    client = GraphBridgeClient(credential=ReusableCredential(), base_url=BASE_URL, max_retries=0)
    tasks = bound_list(client)
    columns_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/columns"
    item_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/items/7"
    create_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/items"
    update_url = f"{item_url}/fields"
    responses.get(columns_url, json={"value": [{"id": "c1", "name": "Title"}]}, status=200)
    responses.get(item_url, json={"id": "7", "fields": {"Title": "Before"}}, status=200)
    responses.post(create_url, json={"id": "8", "fields": {"Title": "New"}}, status=201)
    responses.patch(update_url, json={"Title": "After", "@odata.etag": "etag-2"}, status=200)
    responses.delete(item_url, status=204)

    column_page = tasks.columns.list()
    existing = tasks.items.get("7", fields=("Title",))
    created = tasks.items.create({"Title": "New"})
    updated = tasks.items.update("7", {"Title": "After"}, etag="etag-1")
    deleted = tasks.items.delete("7", etag="etag-2")

    assert column_page.items[0].name == "Title"
    assert existing.fields == {"Title": "Before"}
    assert created.id == "8"
    assert updated.id == "7"
    assert updated.fields["Title"] == "After"
    assert updated.etag == "etag-2"
    assert deleted is None
    assert parse_qs(urlsplit(responses.calls[1].request.url).query) == {
        "$expand": ["fields($select=Title)"]
    }
    assert request_json(2) == {"fields": {"Title": "New"}}
    assert request_json(3) == {"Title": "After"}
    assert responses.calls[3].request.headers["If-Match"] == "etag-1"
    assert responses.calls[4].request.headers["If-Match"] == "etag-2"


@responses.activate
def test_list_and_column_discovery_support_pages_and_direct_lookup() -> None:
    client = GraphBridgeClient(credential=ReusableCredential(), base_url=BASE_URL, max_retries=0)
    site = client.sites.bind({"id": SITE_ID})
    lists_url = f"{BASE_URL}/sites/{SITE_ID}/lists"
    lists_next_link = f"{lists_url}?$skiptoken=list-page-2"
    responses.get(
        lists_url,
        json={"value": [{"id": "list-1", "displayName": "First"}]},
        status=200,
    )
    responses.get(
        lists_url,
        json={
            "value": [{"id": "list-1", "displayName": "First"}],
            "@odata.nextLink": lists_next_link,
        },
        status=200,
    )
    responses.get(
        lists_next_link,
        json={"value": [{"id": "list-2", "displayName": "Second"}]},
        status=200,
    )

    first_page = site.lists.list()
    all_lists = list(site.lists.iter_all())

    assert [item.id for item in first_page.items] == ["list-1"]
    assert [item.id for item in all_lists] == ["list-1", "list-2"]

    tasks = site.lists.bind({"id": LIST_ID})
    columns_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/columns"
    columns_next_link = f"{columns_url}?$skiptoken=column-page-2"
    responses.get(
        columns_url,
        json={
            "value": [{"id": "column-1", "name": "Title"}],
            "@odata.nextLink": columns_next_link,
        },
        status=200,
    )
    responses.get(
        columns_next_link,
        json={"value": [{"id": "column-2", "name": "Status"}]},
        status=200,
    )
    responses.get(
        f"{columns_url}/column-2",
        json={"id": "column-2", "name": "Status"},
        status=200,
    )

    columns = list(tasks.columns.iter_all())
    status_column = tasks.columns.get("column-2")

    assert [column.name for column in columns] == ["Title", "Status"]
    assert status_column.name == "Status"


@responses.activate
def test_new_batch_primitives_report_partial_results() -> None:
    client = GraphBridgeClient(credential=ReusableCredential(), base_url=BASE_URL, max_retries=0)
    tasks = bound_list(client)
    batch_url = f"{BASE_URL}/$batch"
    responses.post(
        batch_url,
        json={
            "responses": [
                {"id": "0", "status": 201, "body": {"id": "10", "fields": {"Title": "One"}}},
                {
                    "id": "1",
                    "status": 400,
                    "body": {"error": {"code": "invalidRequest", "message": "Invalid row"}},
                },
            ]
        },
        status=200,
    )
    responses.post(
        batch_url,
        json={
            "responses": [
                {"id": "0", "status": 204},
                {"id": "1", "status": 412, "body": {"error": {"message": "etag mismatch"}}},
            ]
        },
        status=200,
    )

    created = tasks.items.create_many([{"Title": "One"}, {"Title": "Bad"}])
    deleted = tasks.items.delete_many(["10", "11"])

    assert [item.id for item in created.successes] == ["10"]
    assert [(error.status_code, error.code) for error in created.failures] == [(400, "invalidRequest")]
    assert deleted.successes == ["10"]
    assert deleted.failures[0].status_code == 412
    create_requests = request_json(0)["requests"]  # type: ignore[index]
    assert create_requests[0]["url"] == f"/sites/{SITE_ID}/lists/{LIST_ID}/items"
    assert create_requests[0]["body"] == {"fields": {"Title": "One"}}


def test_binding_and_resource_types_require_no_network() -> None:
    client = GraphBridgeClient(credential=ReusableCredential(), base_url=BASE_URL)
    site = client.sites.bind({"id": SITE_ID})
    tasks = site.lists.bind({"id": LIST_ID})

    assert isinstance(site, SiteResource)
    assert isinstance(site.lists, ListsResource)
    assert isinstance(tasks, SharePointListResource)
    assert isinstance(tasks.items, ListItemsResource)
    assert isinstance(tasks.columns, ColumnsResource)
    assert isinstance(tasks.versions, VersionsResource)


def test_client_validation_representation_and_injected_transport() -> None:
    credential = ReusableCredential()
    transport = GraphTransport(credential, base_url=BASE_URL)
    client = GraphBridgeClient(transport=transport)

    assert client.transport is transport
    assert TOKEN not in repr(client)
    assert "do-not-render" not in repr(client)
    with pytest.raises(TypeError, match="credential is required"):
        GraphBridgeClient()
    with pytest.raises(ValueError, match="cannot be supplied"):
        GraphBridgeClient(credential=credential, transport=transport)


@responses.activate
def test_invalid_resource_shapes_raise_graph_response_error() -> None:
    client = GraphBridgeClient(credential=ReusableCredential(), base_url=BASE_URL, max_retries=0)
    responses.get(f"{BASE_URL}/sites/missing", json={"displayName": "No ID"}, status=200)
    with pytest.raises(GraphInvalidResponseError, match="does not contain an id"):
        client.sites.get("missing")

    tasks = bound_list(client)
    responses.get(
        f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/items",
        json={"value": "not-a-list"},
        status=200,
    )
    with pytest.raises(GraphInvalidResponseError, match="must be a list"):
        tasks.items.list()
