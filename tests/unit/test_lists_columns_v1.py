from __future__ import annotations

import json
from types import SimpleNamespace
from urllib.parse import parse_qs, urlsplit

import pytest
import responses

from graphbridge import GraphBridgeClient
from graphbridge.exceptions import GraphAmbiguousMatchError, GraphNotFoundError
from graphbridge.models import ListInfo, SiteInfo

BASE_URL = "https://graph.schema.invalid/v1.0"
SITE_ID = "schema.invalid,collection,site"
LIST_ID = "schema-list"


class Credential:
    def get_token(self, _scope: str) -> SimpleNamespace:
        return SimpleNamespace(token="schema-token")


def client_and_site() -> tuple[GraphBridgeClient, object]:
    client = GraphBridgeClient(credential=Credential(), base_url=BASE_URL, max_retries=0)
    return client, client.sites.bind(SiteInfo(id=SITE_ID))


def request_json(index: int) -> object:
    body = responses.calls[index].request.body
    if isinstance(body, bytes):
        body = body.decode("utf-8")
    return json.loads(body)


@responses.activate
def test_list_enumeration_id_lookup_and_exact_name_lookup() -> None:
    _client, site = client_and_site()
    lists_url = f"{BASE_URL}/sites/{SITE_ID}/lists"
    next_link = f"{lists_url}?$skiptoken=page2"
    responses.get(
        lists_url,
        json={
            "value": [{"id": "1", "displayName": "First"}],
            "@odata.nextLink": next_link,
        },
        status=200,
    )
    responses.get(next_link, json={"value": [{"id": "2", "displayName": "Tasks"}]}, status=200)
    responses.get(
        f"{lists_url}/2",
        json={"id": "2", "displayName": "Tasks", "eTag": "list-etag"},
        status=200,
    )

    by_name = site.lists.get_by_name("Tasks")
    by_id = site.lists.get_by_id("2")

    assert by_name.id == by_id.id == "2"
    assert by_id.info.etag == "list-etag"
    assert [call.request.url for call in responses.calls[:2]] == [lists_url, next_link]


@responses.activate
def test_duplicate_list_names_raise_an_explicit_error() -> None:
    _client, site = client_and_site()
    lists_url = f"{BASE_URL}/sites/{SITE_ID}/lists"
    responses.get(
        lists_url,
        json={
            "value": [
                {"id": "1", "displayName": "Duplicate"},
                {"id": "2", "displayName": "Duplicate"},
            ]
        },
        status=200,
    )

    with pytest.raises(GraphAmbiguousMatchError) as caught:
        site.lists.get_by_name("Duplicate")

    assert caught.value.candidate_ids == ["1", "2"]


@responses.activate
def test_create_list_supports_template_and_initial_columns() -> None:
    _client, site = client_and_site()
    lists_url = f"{BASE_URL}/sites/{SITE_ID}/lists"
    responses.post(
        lists_url,
        json={
            "id": "created-list",
            "displayName": "Books",
            "list": {"template": "genericList"},
        },
        status=201,
    )

    created = site.lists.create(
        "Books",
        description="Non-destructive test list",
        columns=(
            {"name": "Author", "text": {}},
            {"name": "PageCount", "number": {}},
        ),
    )

    assert created.id == "created-list"
    assert created.info.template == "genericList"
    assert request_json(0) == {
        "displayName": "Books",
        "description": "Non-destructive test list",
        "columns": [
            {"name": "Author", "text": {}},
            {"name": "PageCount", "number": {}},
        ],
        "list": {"template": "genericList"},
    }


@responses.activate
def test_list_metadata_and_expanded_relationships_are_preserved() -> None:
    _client, site = client_and_site()
    list_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}"
    responses.get(
        list_url,
        json={
            "id": LIST_ID,
            "displayName": "Library",
            "description": "Docs",
            "eTag": "etag-list",
            "createdDateTime": "2026-01-01T00:00:00Z",
            "list": {"template": "documentLibrary"},
            "columns": [{"id": "c1", "name": "Title"}],
            "drive": {"id": "drive-1"},
            "unknownMetadata": {"kept": True},
        },
        status=200,
    )

    library = site.lists.get_by_id(
        LIST_ID,
        select=("id", "description", "eTag"),
        expand=("columns", "drive"),
    )

    assert library.info.description == "Docs"
    assert library.info.template == "documentLibrary"
    assert library.relationship("columns")[0]["name"] == "Title"
    assert library.relationship("drive")["id"] == "drive-1"
    assert library.metadata["unknownMetadata"] == {"kept": True}
    assert parse_qs(urlsplit(responses.calls[0].request.url).query) == {
        "$select": ["id,description,eTag"],
        "$expand": ["columns,drive"],
    }


@responses.activate
def test_column_crud_preserves_type_and_unknown_properties() -> None:
    client, site = client_and_site()
    resource = site.lists.bind(ListInfo(id=LIST_ID))
    columns_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/columns"
    responses.get(
        f"{columns_url}/c1",
        json={
            "id": "c1",
            "name": "Title",
            "displayName": "Title",
            "text": {"maxLength": 255},
            "futureProperty": {"kept": True},
        },
        status=200,
    )
    responses.post(
        columns_url,
        json={"id": "c2", "name": "Priority", "choice": {"choices": ["High", "Low"]}},
        status=201,
    )
    responses.patch(
        f"{columns_url}/c2",
        json={"id": "c2", "name": "Priority", "required": True, "choice": {}},
        status=200,
    )
    responses.delete(f"{columns_url}/c2", status=204)

    existing = resource.columns.get_by_id("c1")
    created = resource.columns.create(
        {"name": "Priority", "choice": {"choices": ["High", "Low"]}}
    )
    updated = resource.columns.update("c2", {"required": True})
    deleted = resource.columns.delete("c2")

    assert existing.column_type == "text"
    assert existing.type_properties == {"maxLength": 255}
    assert existing.raw["futureProperty"] == {"kept": True}
    assert created.column_type == updated.column_type == "choice"
    assert deleted is None
    assert client.transport.base_url == BASE_URL


@responses.activate
def test_display_name_mapping_is_cached_and_used_by_item_writes() -> None:
    _client, site = client_and_site()
    resource = site.lists.bind(ListInfo(id=LIST_ID))
    columns_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/columns"
    items_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/items"
    responses.get(
        columns_url,
        json={
            "value": [
                {"id": "c1", "displayName": "Project Name", "name": "Project_x0020_Name"},
                {"id": "c2", "displayName": "Status", "name": "Status"},
            ]
        },
        status=200,
    )
    responses.post(items_url, json={"id": "10"}, status=201)

    assert resource.columns.display_name_map() == {
        "Project Name": "Project_x0020_Name",
        "Status": "Status",
    }
    created = resource.items.create(
        {"Project Name": "Apollo", "Status": "Open"},
        field_names="display",
    )

    assert created.fields == {"Project_x0020_Name": "Apollo", "Status": "Open"}
    assert request_json(1) == {
        "fields": {"Project_x0020_Name": "Apollo", "Status": "Open"}
    }
    assert sum(call.request.method == "GET" for call in responses.calls) == 1


@responses.activate
def test_column_name_lookup_translation_and_duplicate_mapping_errors() -> None:
    _client, site = client_and_site()
    resource = site.lists.bind(ListInfo(id=LIST_ID))
    columns_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/columns"
    responses.get(
        columns_url,
        json={
            "value": [
                {"id": "1", "displayName": "Friendly", "name": "Internal"},
                {"id": "2", "displayName": "Other", "name": "OtherInternal"},
            ]
        },
        status=200,
    )

    assert resource.columns.get_by_name("Internal").id == "1"
    assert resource.columns.get_by_name("Friendly").id == "1"
    assert resource.columns.to_display_fields({"Internal": 1, "Unknown": 2}) == {
        "Friendly": 1,
        "Unknown": 2,
    }
    assert resource.columns.to_internal_fields({"Unknown": 2}, strict=False) == {"Unknown": 2}
    with pytest.raises(GraphNotFoundError):
        resource.columns.get_by_name("Missing")
    with pytest.raises(KeyError, match="unknown"):
        resource.columns.to_internal_fields({"Missing": 1})

    resource.columns.invalidate_schema()
    responses.get(
        columns_url,
        json={
            "value": [
                {"id": "3", "displayName": "Duplicate", "name": "One"},
                {"id": "4", "displayName": "Duplicate", "name": "Two"},
            ]
        },
        status=200,
    )
    with pytest.raises(GraphAmbiguousMatchError):
        resource.columns.display_name_map()
