from __future__ import annotations

import json

import pytest
import responses
from conftest import ACCESS_TOKEN

from graphbridge import GbList


def _json_body(call_index: int) -> object:
    body = responses.calls[call_index].request.body
    if isinstance(body, bytes):
        body = body.decode("utf-8")
    return json.loads(body)


@responses.activate
def test_update_records_successes_failures_and_request_shape(resolved_list: GbList) -> None:
    first_url = f"{resolved_list.list_url}/items/1/fields"
    second_url = f"{resolved_list.list_url}/items/2/fields"
    responses.patch(first_url, json={"id": "1", "Status": "Done"}, status=200)
    responses.patch(second_url, body="simulated conflict", status=409)

    result = resolved_list.update(
        ids=("1", "2"),
        rows=({"Status": "Done"}, {"Status": "Blocked"}),
    )

    assert result == {
        "successes": [
            {
                "id": "1",
                "success": True,
                "updated_row": {"id": "1", "Status": "Done"},
            }
        ],
        "failures": [
            {
                "id": "2",
                "success": False,
                "error": "Error updating: 409 simulated conflict",
            }
        ],
    }
    assert [call.request.method for call in responses.calls] == ["PATCH", "PATCH"]
    assert [call.request.url for call in responses.calls] == [first_url, second_url]
    assert _json_body(0) == {"Status": "Done"}
    assert _json_body(1) == {"Status": "Blocked"}
    assert responses.calls[0].request.headers["Content-Type"] == "application/json"
    assert responses.calls[0].request.headers["Authorization"] == f"Bearer {ACCESS_TOKEN}"


@responses.activate
def test_update_accepts_single_integer_id_and_dict(resolved_list: GbList) -> None:
    responses.patch(
        f"{resolved_list.list_url}/items/7/fields",
        json={"Title": "Updated"},
        status=200,
    )

    result = resolved_list.update(7, {"Title": "Updated"})

    assert result["successes"][0]["id"] is None


@pytest.mark.parametrize(
    ("ids", "rows", "error"),
    [
        (["1", "2"], [{"Title": "Only one"}], ValueError),
        (None, [{"Title": "One"}], TypeError),
        (["1"], None, TypeError),
    ],
)
def test_update_input_errors(resolved_list: GbList, ids: object, rows: object, error: type[Exception]) -> None:
    with pytest.raises(error):
        resolved_list.update(ids=ids, rows=rows)  # type: ignore[arg-type]


@responses.activate
def test_create_records_successes_failures_and_wraps_fields(resolved_list: GbList) -> None:
    create_url = f"{resolved_list.list_url}/items"
    responses.post(create_url, json={"id": "10", "fields": {"Title": "One"}}, status=201)
    responses.post(create_url, body="simulated invalid field", status=400)

    result = resolved_list.create(({"Title": "One"}, {"Title": "Invalid"}))

    assert result == {
        "successes": [
            {
                "id": "10",
                "success": True,
                "item": {"id": "10", "fields": {"Title": "One"}},
            }
        ],
        "failures": [
            {
                "success": False,
                "error": "Error while creating a new item: 400 simulated invalid field",
            }
        ],
    }
    assert _json_body(0) == {"fields": {"Title": "One"}}
    assert _json_body(1) == {"fields": {"Title": "Invalid"}}


@responses.activate
def test_create_accepts_single_dict(resolved_list: GbList) -> None:
    responses.post(f"{resolved_list.list_url}/items", json={"id": "11"}, status=201)

    result = resolved_list.create({"Title": "Single"})

    assert result["successes"][0]["id"] == "11"
    assert _json_body(0) == {"fields": {"Title": "Single"}}


def test_create_rejects_non_iterable(resolved_list: GbList) -> None:
    with pytest.raises(TypeError, match="rows must"):
        resolved_list.create(123)  # type: ignore[arg-type]


@responses.activate
def test_create_legacy_wraps_non_dict_list_items_without_validating_them(resolved_list: GbList) -> None:
    responses.post(f"{resolved_list.list_url}/items", json={"id": "12"}, status=201)

    result = resolved_list.create(["legacy-value"])  # type: ignore[list-item]

    assert result["successes"][0]["id"] == "12"
    assert _json_body(0) == {"fields": "legacy-value"}


@responses.activate
def test_delete_records_successes_failures_and_request_shape(resolved_list: GbList) -> None:
    first_url = f"{resolved_list.list_url}/items/1"
    second_url = f"{resolved_list.list_url}/items/2"
    responses.delete(first_url, status=204)
    responses.delete(second_url, body="simulated forbidden", status=403)

    result = resolved_list.delete(("1", "2"))

    assert result == {
        "successes": [
            {
                "id": "1",
                "completed": True,
                "message": "Item deleted successfully.",
            }
        ],
        "failures": [
            {
                "id": "2",
                "completed": False,
                "error": "Error while deleting: 403 simulated forbidden",
            }
        ],
    }
    assert [call.request.method for call in responses.calls] == ["DELETE", "DELETE"]
    assert responses.calls[0].request.headers["Authorization"] == f"Bearer {ACCESS_TOKEN}"


@pytest.mark.parametrize(("ids", "error"), [([], ValueError), (123, TypeError), (None, TypeError)])
def test_delete_input_errors(resolved_list: GbList, ids: object, error: type[Exception]) -> None:
    with pytest.raises(error):
        resolved_list.delete(ids)  # type: ignore[arg-type]


@responses.activate
def test_delete_accepts_single_string_id(resolved_list: GbList) -> None:
    responses.delete(f"{resolved_list.list_url}/items/7", status=204)

    assert resolved_list.delete("7")["successes"][0]["id"] == "7"
