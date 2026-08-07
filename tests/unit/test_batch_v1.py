from __future__ import annotations

import json
import math
from types import SimpleNamespace
from typing import Any

import pytest
import responses

from graphbridge import GraphBridgeClient
from graphbridge.batch import BatchRequest, batch_payload, execute_batch
from graphbridge.exceptions import GraphInvalidResponseError
from graphbridge.models import ListInfo, SiteInfo

BASE_URL = "https://graph.batch.invalid/v1.0"
SITE_ID = "batch.invalid,collection,site"
LIST_ID = "batch-list"
BATCH_URL = f"{BASE_URL}/$batch"


class Credential:
    def get_token(self, _scope: str) -> SimpleNamespace:
        return SimpleNamespace(token="batch-token")


def tasks() -> object:
    client = GraphBridgeClient(credential=Credential(), base_url=BASE_URL, max_retries=0)
    site = client.sites.bind(SiteInfo(id=SITE_ID))
    return site.lists.bind(ListInfo(id=LIST_ID))


def request_json(index: int) -> dict[str, Any]:
    body = responses.calls[index].request.body
    if isinstance(body, bytes):
        body = body.decode("utf-8")
    return json.loads(body)


@pytest.mark.parametrize("count", [1, 20, 21, 45])
@responses.activate
def test_batch_create_chunks_1_20_21_and_more_in_input_order(count: int) -> None:
    resource = tasks()

    def callback(request: object) -> tuple[int, dict[str, str], str]:
        body = request.body  # type: ignore[attr-defined]
        if isinstance(body, bytes):
            body = body.decode("utf-8")
        requests_body = json.loads(body)["requests"]
        batch_responses = [
            {
                "id": entry["id"],
                "status": 201,
                "body": {"id": f"item-{entry['id']}", "fields": entry["body"]["fields"]},
            }
            for entry in reversed(requests_body)
        ]
        return 200, {"Content-Type": "application/json"}, json.dumps({"responses": batch_responses})

    responses.add_callback(responses.POST, BATCH_URL, callback=callback)
    records = [{"Title": f"Row {index}"} for index in range(count)]

    result = resource.items.create_many(records)

    assert len(responses.calls) == math.ceil(count / 20)
    assert [len(request_json(index)["requests"]) for index in range(len(responses.calls))] == [
        min(20, count - start) for start in range(0, count, 20)
    ]
    assert [item.id for item in result.successes] == [f"item-{index}" for index in range(count)]
    assert [entry.input_index for entry in result.results] == list(range(count))
    assert [entry.request_id for entry in result.results] == [str(index) for index in range(count)]


@responses.activate
def test_partial_batch_response_keeps_safe_input_correlation() -> None:
    resource = tasks()
    responses.post(
        BATCH_URL,
        json={
            "responses": [
                {
                    "id": "1",
                    "status": 400,
                    "body": {"error": {"code": "invalidRequest", "message": "bad row"}},
                },
                {"id": "0", "status": 201, "body": {"id": "100"}},
            ]
        },
        status=200,
    )

    result = resource.items.create_many([{"Title": "Good"}, {"Title": "Bad"}])

    assert [item.id for item in result.successes] == ["100"]
    assert result.failures[0].inner_error["input_index"] == 1
    assert [entry.status_code for entry in result.results] == [201, 400]
    assert result.results[0].value is result.successes[0]
    assert result.results[1].error is result.failures[0]


@responses.activate
def test_batch_retries_only_throttled_subrequests_and_caps_retry_after() -> None:
    resource = tasks()
    resource.transport.max_retry_delay = 2.0
    sleeps: list[float] = []
    responses.post(
        BATCH_URL,
        json={
            "responses": [
                {"id": "0", "status": 201, "body": {"id": "10"}},
                {
                    "id": "1",
                    "status": 429,
                    "headers": {"Retry-After": "999"},
                    "body": {"error": {"code": "tooManyRequests", "message": "slow down"}},
                },
            ]
        },
        status=200,
    )
    responses.post(
        BATCH_URL,
        json={"responses": [{"id": "1", "status": 201, "body": {"id": "11"}}]},
        status=200,
    )

    result = resource.items.create_many(
        [{"Title": "First"}, {"Title": "Second"}],
        max_attempts=3,
        sleep=sleeps.append,
    )

    assert sleeps == [2.0]
    assert [entry["id"] for entry in request_json(0)["requests"]] == ["0", "1"]
    assert [entry["id"] for entry in request_json(1)["requests"]] == ["1"]
    assert [item.id for item in result.successes] == ["10", "11"]
    assert [entry.attempts for entry in result.results] == [1, 2]


@responses.activate
def test_batch_retry_budget_prevents_infinite_loop() -> None:
    resource = tasks()
    for _ in range(2):
        responses.post(
            BATCH_URL,
            json={
                "responses": [
                    {
                        "id": "0",
                        "status": 429,
                        "headers": {"retry-after": "0"},
                        "body": {"error": {"message": "still throttled"}},
                    }
                ]
            },
            status=200,
        )

    result = resource.items.delete_many(["10"], max_attempts=2, sleep=lambda _delay: None)

    assert len(responses.calls) == 2
    assert result.successes == []
    assert result.failures[0].status_code == 429
    assert result.results[0].attempts == 2


@responses.activate
def test_update_many_and_delete_many_apply_per_item_etags() -> None:
    resource = tasks()
    responses.post(
        BATCH_URL,
        json={
            "responses": [
                {"id": "0", "status": 200, "body": {"Title": "After", "@odata.etag": "new"}},
                {
                    "id": "1",
                    "status": 412,
                    "body": {"error": {"code": "preconditionFailed", "message": "stale"}},
                },
            ]
        },
        status=200,
    )
    responses.post(
        BATCH_URL,
        json={"responses": [{"id": "0", "status": 204}, {"id": "1", "status": 204}]},
        status=200,
    )

    updated = resource.items.update_many(
        [
            ("10", {"Title": "After"}, "etag-10"),
            {"id": "11", "fields": {"Title": "Stale"}, "etag": "etag-11"},
        ]
    )
    deleted = resource.items.delete_many(["10", "11"], etags={"10": "new", "11": "*"})

    update_requests = request_json(0)["requests"]
    delete_requests = request_json(1)["requests"]
    assert [entry["method"] for entry in update_requests] == ["PATCH", "PATCH"]
    assert [entry["headers"]["If-Match"] for entry in update_requests] == [
        "etag-10",
        "etag-11",
    ]
    assert updated.successes[0].id == "10"
    assert updated.successes[0].etag == "new"
    assert updated.failures[0].status_code == 412
    assert [entry["headers"]["If-Match"] for entry in delete_requests] == ["new", "*"]
    assert deleted.successes == ["10", "11"]


def test_generic_batch_validates_v1_urls_ids_and_retry_configuration() -> None:
    request = BatchRequest(id="A", method="patch", url="/items/1", body={"x": 1})
    assert request.to_payload()["headers"] == {"Content-Type": "application/json"}
    with pytest.raises(ValueError, match="v1.0"):
        BatchRequest(id="1", method="GET", url="/beta/users").to_payload()
    with pytest.raises(ValueError, match="v1.0"):
        BatchRequest(id="1", method="GET", url="//attacker.invalid/users").to_payload()
    with pytest.raises(ValueError, match="method"):
        BatchRequest(id="1", method="TRACE", url="/users").to_payload()
    with pytest.raises(ValueError, match="unique"):
        batch_payload([request, BatchRequest(id="a", method="GET", url="/items/2")])
    secret = "batch-bearer-secret"
    unsafe = BatchRequest(
        id="2",
        method="GET",
        url="/items/2",
        headers={"Authorization": f"Bearer {secret}"},
        body={"secret": secret},
    )
    assert secret not in repr(unsafe)
    with pytest.raises(ValueError, match="Authorization"):
        unsafe.to_payload()
    with pytest.raises(ValueError, match="at least one"):
        execute_batch(None, [], max_attempts=0)  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="non-negative"):
        execute_batch(None, [], backoff_factor=-1)  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="cannot exceed"):
        execute_batch(None, [], max_attempts=12)  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="finite"):
        execute_batch(None, [], max_retry_delay=float("inf"))  # type: ignore[arg-type]


@responses.activate
def test_generic_batch_rejects_uncorrelatable_subresponses() -> None:
    resource = tasks()
    responses.post(
        BATCH_URL,
        json={"responses": [{"id": "unknown", "status": 201, "body": {"id": "1"}}]},
        status=200,
    )

    with pytest.raises(GraphInvalidResponseError, match="unknown"):
        resource.items.create_many([{"Title": "One"}])
