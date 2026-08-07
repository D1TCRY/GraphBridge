from __future__ import annotations

import json

import pytest
import responses
from conftest import ACCESS_TOKEN, LIST_ID, SITE_ID

from graphbridge import GbList

BATCH_URL = "https://graph.microsoft.com/v1.0/$batch"


def _body(index: int) -> dict[str, object]:
    body = responses.calls[index].request.body
    if isinstance(body, bytes):
        body = body.decode("utf-8")
    return json.loads(body)


@responses.activate
def test_create_many_chunks_requests_and_preserves_subresponses(resolved_list: GbList) -> None:
    rows = [{"Title": f"Item {index}"} for index in range(21)]
    responses.post(
        BATCH_URL,
        json={
            "responses": [
                {"id": "1-1", "status": 201, "body": {"id": "101"}},
                {"id": "1-2", "status": 400, "body": {"error": "invalid"}},
            ]
        },
        status=200,
    )
    responses.post(
        BATCH_URL,
        json={"responses": [{"id": "2-1", "status": 201, "body": {"id": "121"}}]},
        status=200,
    )

    result = resolved_list.create_many(rows)

    assert result == {
        "successes": [
            {"id": "101", "status": 201, "item": {"id": "101"}},
            {"id": "121", "status": 201, "item": {"id": "121"}},
        ],
        "failures": [
            {"id": "1-2", "status": 400, "error": {"error": "invalid"}}
        ],
    }
    assert len(responses.calls) == 2
    first_requests = _body(0)["requests"]
    second_requests = _body(1)["requests"]
    assert isinstance(first_requests, list) and len(first_requests) == 20
    assert isinstance(second_requests, list) and len(second_requests) == 1
    assert first_requests[0] == {
        "id": "1-1",
        "method": "POST",
        "url": f"/sites/{SITE_ID}/lists/{LIST_ID}/items",
        "headers": {"Content-Type": "application/json"},
        "body": {"fields": {"Title": "Item 0"}},
    }
    assert responses.calls[0].request.headers["Authorization"] == f"Bearer {ACCESS_TOKEN}"
    assert responses.calls[0].request.headers["Content-Type"] == "application/json"


@responses.activate
def test_create_many_records_outer_batch_failure(resolved_list: GbList) -> None:
    responses.post(BATCH_URL, body="simulated batch failure", status=503)

    assert resolved_list.create_many([{"Title": "One"}]) == {
        "successes": [],
        "failures": [{"batch": 1, "error": "503 simulated batch failure"}],
    }


@pytest.mark.parametrize("rows", [None, {"Title": "One"}, ["not-a-dict"]])
def test_create_many_validates_rows(resolved_list: GbList, rows: object) -> None:
    with pytest.raises(TypeError, match="lista di dict"):
        resolved_list.create_many(rows)  # type: ignore[arg-type]


@responses.activate
def test_delete_many_maps_ids_and_if_match_headers(resolved_list: GbList) -> None:
    responses.post(
        BATCH_URL,
        json={
            "responses": [
                {"id": "1-1", "status": 204},
                {"id": "1-2", "status": 412, "body": {"error": "etag mismatch"}},
            ]
        },
        status=200,
    )

    result = resolved_list.delete_many((10, 11), if_match="*")

    assert result == {
        "successes": [{"id": "10", "status": 204}],
        "failures": [
            {"id": "11", "status": 412, "error": {"error": "etag mismatch"}}
        ],
    }
    requests_body = _body(0)["requests"]
    assert requests_body == [
        {
            "id": "1-1",
            "method": "DELETE",
            "url": f"/sites/{SITE_ID}/lists/{LIST_ID}/items/10",
            "headers": {"If-Match": "*"},
        },
        {
            "id": "1-2",
            "method": "DELETE",
            "url": f"/sites/{SITE_ID}/lists/{LIST_ID}/items/11",
            "headers": {"If-Match": "*"},
        },
    ]


@responses.activate
def test_delete_many_accepts_single_integer_without_if_match(resolved_list: GbList) -> None:
    responses.post(BATCH_URL, json={"responses": [{"id": "1-1", "status": 204}]}, status=200)

    assert resolved_list.delete_many(7) == {
        "successes": [{"id": "7", "status": 204}],
        "failures": [],
    }
    assert "headers" not in _body(0)["requests"][0]


@responses.activate
def test_delete_many_records_outer_batch_failure(resolved_list: GbList) -> None:
    responses.post(BATCH_URL, body="simulated unavailable", status=503)

    assert resolved_list.delete_many(["1"]) == {
        "successes": [],
        "failures": [{"batch": 1, "error": "503 simulated unavailable"}],
    }


def test_delete_many_rejects_non_iterable_and_accepts_empty_list(resolved_list: GbList) -> None:
    with pytest.raises(TypeError, match="iterabile"):
        resolved_list.delete_many(None)

    assert resolved_list.delete_many([]) == {"successes": [], "failures": []}
