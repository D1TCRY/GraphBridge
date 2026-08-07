from __future__ import annotations

from types import SimpleNamespace

import pytest
import responses

from graphbridge import GraphBridgeClient
from graphbridge.exceptions import DeltaResetRequiredError, GraphGoneError
from graphbridge.models import ListInfo, SiteInfo

BASE_URL = "https://graph.delta.invalid/v1.0"
SITE_ID = "delta.invalid,collection,site"
LIST_ID = "delta-list"


class Credential:
    def get_token(self, _scope: str) -> SimpleNamespace:
        return SimpleNamespace(token="delta-token")


def tasks() -> object:
    client = GraphBridgeClient(credential=Credential(), base_url=BASE_URL, max_retries=0)
    site = client.sites.bind(SiteInfo(id=SITE_ID))
    return site.lists.bind(ListInfo(id=LIST_ID))


def delta_url() -> str:
    return f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/items/delta"


def versions_url(item_id: str = "7") -> str:
    return f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/items/{item_id}/versions"


@responses.activate
def test_delta_one_page_classifies_created_modified_and_deleted() -> None:
    resource = tasks()
    cursor = f"{delta_url()}?token=opaque%2Bcursor"
    responses.get(
        delta_url(),
        json={
            "value": [
                {"id": "1", "eTag": "etag-2", "fields": {"Title": "Changed"}},
                {"id": "2", "fields": {"Title": "Created"}},
                {"id": "3", "deleted": {"state": "deleted"}},
            ],
            "@odata.deltaLink": cursor,
        },
        status=200,
    )

    delta = resource.items.delta(known_ids={"1", "3"}, fields=("Title",), top=50)

    assert [item.id for item in delta.created] == ["2"]
    assert [(item.id, item.etag) for item in delta.modified] == [("1", "etag-2")]
    assert [(item.id, item.state) for item in delta.deleted] == [("3", "deleted")]
    assert delta.unclassified == []
    assert delta.delta_link == cursor
    assert delta.cursor == cursor
    assert delta.pages == 1


@responses.activate
def test_delta_follows_multiple_pages_verbatim_and_keeps_last_occurrence() -> None:
    resource = tasks()
    first = delta_url()
    second = f"{first}?$skiptoken=opaque%2Bpage"
    cursor = f"{first}?token=opaque%2Bfinal"
    responses.get(
        first,
        json={
            "value": [{"id": "1", "fields": {"Title": "Intermediate"}}],
            "@odata.nextLink": second,
        },
        status=200,
    )
    responses.get(
        second,
        json={
            "value": [
                {"id": "1", "fields": {"Title": "Final"}},
                {"id": "2", "deleted": {"state": "deleted"}},
            ],
            "@odata.deltaLink": cursor,
        },
        status=200,
    )

    delta = resource.items.delta(link=first, known_ids={"1", "2"})

    assert [(item.id, item.fields["Title"]) for item in delta.modified] == [("1", "Final")]
    assert [item.id for item in delta.deleted] == ["2"]
    assert delta.pages == 2
    assert [call.request.url for call in responses.calls] == [first, second]
    assert delta.delta_link == cursor


@responses.activate
def test_delta_without_known_state_does_not_guess_create_vs_modify() -> None:
    resource = tasks()
    cursor = f"{delta_url()}?token=next"
    responses.get(
        delta_url(),
        json={"value": [{"id": "1", "fields": {"Title": "Upsert"}}], "@odata.deltaLink": cursor},
        status=200,
    )

    delta = resource.items.delta()

    assert delta.created == delta.modified == []
    assert [item.id for item in delta.unclassified] == ["1"]


@responses.activate
def test_delta_opaque_cursor_is_not_modified() -> None:
    resource = tasks()
    cursor = f"{delta_url()}?token=A%2BB%2Fopaque%3D"
    next_cursor = f"{delta_url()}?token=next%2Bopaque"
    responses.get(cursor, json={"value": [], "@odata.deltaLink": next_cursor}, status=200)

    result = resource.items.delta(link=cursor, known_ids=set())

    assert responses.calls[0].request.url == cursor
    assert result.delta_link == next_cursor
    with pytest.raises(ValueError, match="cannot be combined"):
        resource.items.delta(link=cursor, token="latest")


@responses.activate
def test_delta_410_requires_explicit_rebuild_decision() -> None:
    resource = tasks()
    expired = f"{delta_url()}?token=expired"
    restart = f"{delta_url()}?token=fresh-enumeration"
    responses.get(
        expired,
        json={
            "error": {
                "code": "resyncChangesUploadDifferences",
                "message": "token expired",
            }
        },
        headers={"Location": restart},
        status=410,
    )

    with pytest.raises(DeltaResetRequiredError) as caught:
        resource.items.delta(link=expired)

    assert caught.value.strategy == "resyncChangesUploadDifferences"
    assert caught.value.restart_link == restart
    assert len(responses.calls) == 1


@responses.activate
def test_generic_410_is_not_misreported_as_a_delta_reset() -> None:
    resource = tasks()
    expired = f"{delta_url()}?token=gone"
    responses.get(
        expired,
        json={"error": {"code": "genericGone", "message": "gone"}},
        status=410,
    )

    with pytest.raises(GraphGoneError) as caught:
        resource.items.delta(link=expired)
    assert caught.value.error.code == "genericGone"


@responses.activate
def test_versions_are_paginated_and_preserve_version_metadata() -> None:
    resource = tasks()
    second = f"{versions_url()}?$skiptoken=second"
    responses.get(
        versions_url(),
        json={
            "value": [
                {
                    "id": "3.0",
                    "lastModifiedDateTime": "2026-01-03T00:00:00Z",
                    "lastModifiedBy": {"user": {"displayName": "Anonymous"}},
                    "published": {"level": "published"},
                    "fields": {"Title": "Third"},
                }
            ],
            "@odata.nextLink": second,
        },
        status=200,
    )
    responses.get(
        second,
        json={"value": [{"id": "2.0", "fields": {"Title": "Second"}}]},
        status=200,
    )

    versions = resource.items.versions("7")

    assert [version.id for version in versions] == ["3.0", "2.0"]
    assert versions[0].fields == {"Title": "Third"}
    assert versions[0].last_modified_by["user"]["displayName"] == "Anonymous"
    assert [call.request.url for call in responses.calls] == [versions_url(), second]


@responses.activate
def test_restore_version_uses_stable_restore_version_action() -> None:
    resource = tasks()
    restore_url = f"{versions_url()}/2.0/restoreVersion"
    responses.post(restore_url, status=204)

    assert resource.items.restore_version("7", "2.0") is None
    assert responses.calls[0].request.method == "POST"
    assert responses.calls[0].request.body is None
