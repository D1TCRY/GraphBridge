from __future__ import annotations

import json
from types import SimpleNamespace
from typing import Any

import pytest
import responses

from graphbridge import GraphBridgeClient
from graphbridge.exceptions import SyncDuplicateKeyError, SyncMissingKeyError
from graphbridge.models import ListInfo, SiteInfo, SyncOperation

BASE_URL = "https://graph.sync.invalid/v1.0"
SITE_ID = "sync.invalid,collection,site"
LIST_ID = "sync-list"


class Credential:
    def get_token(self, _scope: str) -> SimpleNamespace:
        return SimpleNamespace(token="sync-token")


def tasks() -> object:
    client = GraphBridgeClient(credential=Credential(), base_url=BASE_URL, max_retries=0)
    site = client.sites.bind(SiteInfo(id=SITE_ID))
    return site.lists.bind(ListInfo(id=LIST_ID, display_name="Sync"))


def items_url() -> str:
    return f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/items"


def batch_url() -> str:
    return f"{BASE_URL}/$batch"


def request_json(index: int) -> dict[str, Any]:
    body = responses.calls[index].request.body
    if isinstance(body, bytes):
        body = body.decode("utf-8")
    assert isinstance(body, str)
    value = json.loads(body)
    assert isinstance(value, dict)
    return value


def echo_success_callback(request: Any) -> tuple[int, dict[str, str], str]:
    body = request.body.decode("utf-8") if isinstance(request.body, bytes) else request.body
    payload = json.loads(body)
    batch_responses = []
    for entry in payload["requests"]:
        method = entry["method"]
        if method == "POST":
            response_body: Any = {
                "id": f"created-{entry['id']}",
                "fields": entry["body"]["fields"],
            }
            status = 201
        elif method == "PATCH":
            response_body = entry["body"]
            status = 200
        else:
            response_body = None
            status = 204
        batch_responses.append(
            {"id": entry["id"], "status": status, "body": response_body}
        )
    return 200, {"Content-Type": "application/json"}, json.dumps({"responses": batch_responses})


@responses.activate
def test_sync_plan_without_changes_is_serializable_and_side_effect_free() -> None:
    resource = tasks()
    responses.get(
        items_url(),
        json={
            "value": [
                {
                    "id": "1",
                    "eTag": 'W/"1"',
                    "fields": {"ExternalId": "A", "Title": "Same", "ServerOnly": 1},
                }
            ]
        },
        status=200,
    )

    plan = resource.sync.plan(
        rows=[{"ExternalId": "A", "Title": "Same"}],
        key_field="ExternalId",
    )

    assert plan.operation_count == 0
    assert len(plan.unchanged) == 1
    assert plan.unchanged[0].etag == 'W/"1"'
    assert plan.to_dict()["unchanged"][0]["reason"] == "source fields already match SharePoint"
    assert [call.request.method for call in responses.calls] == ["GET"]


@responses.activate
def test_sync_plan_classifies_create_update_unchanged_and_prune_modes() -> None:
    remote = {
        "value": [
            {"id": "1", "eTag": "etag-a", "fields": {"ExternalId": "A", "Title": "Old"}},
            {"id": "2", "fields": {"ExternalId": "B", "Title": "Same"}},
            {"id": "3", "eTag": "etag-c", "fields": {"ExternalId": "C", "Title": "Remote"}},
        ]
    }
    rows = [
        {"ExternalId": "A", "Title": "New"},
        {"ExternalId": "B", "Title": "Same"},
        {"ExternalId": "D", "Title": "Create"},
    ]
    responses.get(items_url(), json=remote, status=200)
    resource = tasks()
    safe = resource.sync.plan(rows=rows, key_field="ExternalId")

    assert [operation.key for operation in safe.creates] == ["D"]
    assert [operation.key for operation in safe.updates] == ["A"]
    assert safe.updates[0].fields == {"Title": "New"}
    assert safe.updates[0].etag == "etag-a"
    assert safe.deletes == []
    assert [operation.key for operation in safe.unchanged] == ["B", "C"]
    assert "prune is disabled" in safe.unchanged[-1].reason

    responses.get(items_url(), json=remote, status=200)
    pruned = resource.sync.plan(rows=rows, key_field="ExternalId", prune=True)
    assert [operation.key for operation in pruned.deletes] == ["C"]
    assert pruned.deletes[0].etag == "etag-c"
    assert [operation.key for operation in pruned.unchanged] == ["B"]


@responses.activate
def test_sync_dry_run_never_applies_mutations() -> None:
    resource = tasks()
    responses.get(items_url(), json={"value": []}, status=200)

    plan = resource.sync.plan(
        rows=[{"ExternalId": "A", "Title": "New"}],
        key_field="ExternalId",
        dry_run=True,
    )
    result = resource.sync.apply(plan)

    assert result.applied is False
    assert result.dry_run is True
    assert result.created == []
    assert [call.request.method for call in responses.calls] == ["GET"]


@responses.activate
def test_sync_validates_missing_and_duplicate_keys_locally_and_remotely() -> None:
    resource = tasks()
    with pytest.raises(SyncMissingKeyError) as missing:
        resource.sync.plan(rows=[{"Title": "Missing"}], key_field="ExternalId")
    assert missing.value.locations == ["source[0]"]

    with pytest.raises(SyncDuplicateKeyError) as duplicate:
        resource.sync.plan(
            rows=[{"ExternalId": "A"}, {"ExternalId": "A"}],
            key_field="ExternalId",
        )
    assert duplicate.value.duplicates["A"] == ["source[0]", "source[1]"]
    assert len(responses.calls) == 0

    responses.get(
        items_url(),
        json={"value": [{"id": "1", "fields": {"Title": "No key"}}]},
        status=200,
    )
    with pytest.raises(SyncMissingKeyError) as remote_missing:
        resource.sync.plan(rows=[], key_field="ExternalId")
    assert "remote[0]" in remote_missing.value.locations[0]


@responses.activate
def test_empty_source_is_non_destructive_unless_prune_is_explicit() -> None:
    remote = {"value": [{"id": "1", "eTag": "etag", "fields": {"ExternalId": "A"}}]}
    resource = tasks()
    responses.get(items_url(), json=remote, status=200)
    safe = resource.sync.plan(rows=[], key_field="ExternalId")
    assert safe.deletes == []
    assert [operation.key for operation in safe.unchanged] == ["A"]

    responses.get(items_url(), json=remote, status=200)
    pruned = resource.sync.plan(rows=[], key_field="ExternalId", prune=True)
    assert [operation.item_id for operation in pruned.deletes] == ["1"]


@responses.activate
def test_sync_uses_authoritative_schema_mapping_for_display_names() -> None:
    resource = tasks()
    columns_url = f"{BASE_URL}/sites/{SITE_ID}/lists/{LIST_ID}/columns"
    responses.get(
        columns_url,
        json={
            "value": [
                {"id": "c1", "displayName": "External ID", "name": "External_x0020_ID"},
                {"id": "c2", "displayName": "Task title", "name": "Title"},
            ]
        },
        status=200,
    )
    responses.get(items_url(), json={"value": []}, status=200)

    plan = resource.sync.plan(
        rows=[{"External ID": "A", "Task title": "Mapped"}],
        key_field="External ID",
        field_names="display",
    )

    assert plan.creates[0].fields == {
        "External_x0020_ID": "A",
        "Title": "Mapped",
    }
    assert [call.request.method for call in responses.calls] == ["GET", "GET"]


@responses.activate
def test_sync_apply_uses_patch_etag_and_reports_412_without_hiding_partial_success() -> None:
    resource = tasks()
    responses.get(
        items_url(),
        json={
            "value": [
                {"id": "1", "eTag": 'W/"old"', "fields": {"ExternalId": "A", "Title": "Old"}},
                {"id": "2", "fields": {"ExternalId": "B", "Title": "Same"}},
            ]
        },
        status=200,
    )
    plan = resource.sync.plan(
        rows=[
            {"ExternalId": "A", "Title": "New"},
            {"ExternalId": "B", "Title": "Same"},
            {"ExternalId": "C", "Title": "Create"},
        ],
        key_field="ExternalId",
    )
    responses.post(
        batch_url(),
        json={
            "responses": [
                {"id": "0", "status": 201, "body": {"id": "3", "fields": {"ExternalId": "C"}}}
            ]
        },
        status=200,
    )
    responses.post(
        batch_url(),
        json={
            "responses": [
                {
                    "id": "0",
                    "status": 412,
                    "body": {"error": {"code": "preconditionFailed", "message": "etag mismatch"}},
                }
            ]
        },
        status=200,
    )

    result = resource.sync.apply(plan)

    assert [item.id for item in result.created] == ["3"]
    assert result.updated == []
    assert [error.status_code for error in result.failures] == [412]
    update_request = request_json(2)["requests"][0]
    assert update_request["method"] == "PATCH"
    assert update_request["headers"]["If-Match"] == 'W/"old"'
    assert update_request["body"] == {"Title": "New"}
    assert [operation.key for operation in result.unchanged] == ["B"]


@responses.activate
def test_sync_retry_replays_only_failed_operations() -> None:
    resource = tasks()
    responses.get(
        items_url(),
        json={
            "value": [
                {"id": "1", "fields": {"ExternalId": "A", "Title": "Old"}},
                {"id": "2", "fields": {"ExternalId": "B", "Title": "Old"}},
            ]
        },
        status=200,
    )
    plan = resource.sync.plan(
        rows=[
            {"ExternalId": "A", "Title": "New A"},
            {"ExternalId": "B", "Title": "New B"},
        ],
        key_field="ExternalId",
    )
    responses.post(
        batch_url(),
        json={
            "responses": [
                {"id": "1", "status": 400, "body": {"error": {"code": "bad", "message": "failed B"}}},
                {"id": "0", "status": 200, "body": {"Title": "New A"}},
            ]
        },
        status=200,
    )
    first = resource.sync.apply(plan)
    responses.post(
        batch_url(),
        json={"responses": [{"id": "0", "status": 200, "body": {"Title": "New B"}}]},
        status=200,
    )

    retried = resource.sync.retry(first)

    assert [item.id for item in first.updated] == ["1"]
    assert [result.operation.key for result in first.results if not result.succeeded] == ["B"]
    assert [item.id for item in retried.updated] == ["2"]
    retry_requests = request_json(2)["requests"]
    assert len(retry_requests) == 1
    assert retry_requests[0]["url"].endswith("/items/2/fields")


@responses.activate
def test_failed_create_defers_prune_and_never_deletes_before_create() -> None:
    resource = tasks()
    responses.get(
        items_url(),
        json={"value": [{"id": "1", "fields": {"ExternalId": "OLD"}}]},
        status=200,
    )
    plan = resource.sync.plan(
        rows=[{"ExternalId": "NEW"}], key_field="ExternalId", prune=True
    )
    responses.post(
        batch_url(),
        json={"responses": [{"id": "0", "status": 400, "body": {"error": {"code": "bad", "message": "no create"}}}]},
        status=200,
    )

    result = resource.sync.apply(plan)

    assert [call.request.method for call in responses.calls] == ["GET", "POST"]
    assert result.deleted == []
    assert [error.code for error in result.failures] == ["bad", "deleteDeferred"]
    assert [result.operation.operation for result in result.results if result.deferred] == ["delete"]
    assert [operation.operation for operation in result.retry_plan().deletes] == ["delete"]


@responses.activate
def test_large_dataset_is_split_into_safe_batches() -> None:
    resource = tasks()
    rows = [{"ExternalId": f"K-{index}", "Title": str(index)} for index in range(45)]
    responses.get(items_url(), json={"value": []}, status=200)
    plan = resource.sync.plan(rows=rows, key_field="ExternalId")
    responses.add_callback(responses.POST, batch_url(), callback=echo_success_callback)
    responses.add_callback(responses.POST, batch_url(), callback=echo_success_callback)
    responses.add_callback(responses.POST, batch_url(), callback=echo_success_callback)

    result = resource.sync.apply(plan)

    assert len(result.created) == 45
    payload_sizes = [len(request_json(index)["requests"]) for index in (1, 2, 3)]
    assert payload_sizes == [20, 20, 5]


def test_sync_rejects_non_typed_handwritten_plan_entries() -> None:
    resource = tasks()
    from graphbridge.models import SyncPlan

    with pytest.raises(TypeError, match="SyncOperation"):
        resource.sync.apply(SyncPlan(creates=[{"ExternalId": "A"}]))
    assert isinstance(
        SyncOperation(operation="create", key="A", fields={"ExternalId": "A"}),
        SyncOperation,
    )
