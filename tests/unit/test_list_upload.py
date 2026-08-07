from __future__ import annotations

import pytest
import responses

from graphbridge import GbList


def _items_url(gb_list: GbList) -> str:
    return f"{gb_list.list_url}/items?expand=fields&$top=200"


@responses.activate
def test_upload_reads_existing_items_once_then_updates_and_creates(resolved_list: GbList) -> None:
    responses.get(
        _items_url(resolved_list),
        json={"value": [{"id": "1", "fields": {"Title": "Old"}}]},
        status=200,
    )
    responses.patch(f"{resolved_list.list_url}/items/1/fields", json={"id": "1", "Title": "Updated"}, status=200)
    responses.post(f"{resolved_list.list_url}/items", json={"id": "9", "fields": {"Title": "New"}}, status=201)

    with pytest.warns(DeprecationWarning, match=r"sync\.plan\(\)") as caught:
        result = resolved_list.upload(
            ids=["1", "external-2"],
            rows=[{"Title": "Updated"}, {"Title": "New"}],
        )
    assert caught[0].filename.endswith("test_list_upload.py")

    assert result == {
        "delete_results": {"successes": None, "failures": None},
        "force_results": {
            "replaced": {"successes": [], "failures": []},
            "updated": {
                "successes": [{"id": "1", "row": {"Title": "Updated"}}],
                "failures": [],
            },
            "created": {
                "successes": [
                    {
                        "id": "external-2",
                        "row": {"Title": "New"},
                        "new_id": "9",
                    }
                ],
                "failures": [],
            },
        },
    }
    # Creates are completed before updates and any optional prune phase.
    assert [call.request.method for call in responses.calls] == ["GET", "POST", "PATCH"]
    assert sum(call.request.method == "GET" for call in responses.calls) == 1


@responses.activate
def test_upload_delete_removes_ids_not_present_in_source(resolved_list: GbList) -> None:
    responses.get(
        _items_url(resolved_list),
        json={
            "value": [
                {"id": "1", "fields": {"Title": "Keep"}},
                {"id": "2", "fields": {"Title": "Remove"}},
            ]
        },
        status=200,
    )
    responses.delete(f"{resolved_list.list_url}/items/2", status=204)
    responses.patch(f"{resolved_list.list_url}/items/1/fields", json={"id": "1"}, status=200)

    result = resolved_list.upload(
        ids=["1"],
        rows=[{"Title": "Keep updated"}],
        delete=True,
    )

    assert result["delete_results"]["successes"] == [
        {"id": "2", "completed": True, "message": "Item deleted successfully."}
    ]
    assert [call.request.method for call in responses.calls] == ["GET", "PATCH", "DELETE"]


@responses.activate
def test_upload_delete_reports_empty_cleanup_when_source_already_matches(resolved_list: GbList) -> None:
    responses.get(
        _items_url(resolved_list),
        json={"value": [{"id": "1", "fields": {"Title": "Keep"}}]},
        status=200,
    )
    responses.patch(f"{resolved_list.list_url}/items/1/fields", json={"id": "1"}, status=200)

    result = resolved_list.upload("1", {"Title": "Keep"}, delete=True)

    assert result["delete_results"] == {"successes": [], "failures": []}
    assert [call.request.method for call in responses.calls] == ["GET"]


@responses.activate
def test_upload_force_uses_patch_and_never_delete_then_create(resolved_list: GbList) -> None:
    responses.get(
        _items_url(resolved_list),
        json={"value": [{"id": "1", "fields": {"Title": "Existing"}}]},
        status=200,
    )
    responses.patch(
        f"{resolved_list.list_url}/items/1/fields",
        json={"id": "1", "Title": "Replacement"},
        status=200,
    )

    result = resolved_list.upload(
        ids="1",
        rows={"Title": "Replacement"},
        force=True,
    )

    assert [call.request.method for call in responses.calls] == ["GET", "PATCH"]
    assert result["force_results"]["replaced"]["successes"] == [
        {"id": "1", "row": {"Title": "Replacement"}, "new_id": "1"}
    ]


@responses.activate
def test_upload_force_preserves_legacy_result_shape_without_changing_item_id(resolved_list: GbList) -> None:
    responses.get(
        _items_url(resolved_list),
        json={"value": [{"id": "1", "fields": {"Title": "Existing"}}]},
        status=200,
    )
    responses.patch(f"{resolved_list.list_url}/items/1/fields", json={"id": "1"}, status=200)

    result = resolved_list.upload("1", {"Title": "Replacement"}, force=True)

    assert result["force_results"]["replaced"]["successes"] == [
        {"id": "1", "row": {"Title": "Replacement"}, "new_id": "1"}
    ]
    assert [call.request.method for call in responses.calls] == ["GET", "PATCH"]


@responses.activate
def test_upload_force_reports_patch_failure_without_deleting(resolved_list: GbList) -> None:
    responses.get(
        _items_url(resolved_list),
        json={"value": [{"id": "1", "fields": {"Title": "Existing"}}]},
        status=200,
    )
    responses.patch(
        f"{resolved_list.list_url}/items/1/fields",
        body="simulated patch failure",
        status=403,
    )

    result = resolved_list.upload("1", {"Title": "Replacement"}, force=True)

    assert [call.request.method for call in responses.calls] == ["GET", "PATCH"]
    assert result["force_results"]["replaced"]["failures"] == [
        {
            "id": "1",
            "row": {"Title": "Replacement"},
            "error": "403 Microsoft Graph returned HTTP 403",
        }
    ]


@responses.activate
def test_upload_reports_update_and_create_failures(resolved_list: GbList) -> None:
    responses.get(_items_url(resolved_list), json={"value": [{"id": "1", "fields": {}}]}, status=200)
    responses.patch(f"{resolved_list.list_url}/items/1/fields", body="update failed", status=400)
    responses.post(f"{resolved_list.list_url}/items", body="create failed", status=400)

    result = resolved_list.upload(["1", "2"], [{"Title": "A"}, {"Title": "B"}])

    assert result["force_results"]["updated"]["failures"][0]["id"] == "1"
    assert result["force_results"]["created"]["failures"][0]["id"] == "2"


@pytest.mark.parametrize(
    ("ids", "rows", "force", "delete", "error"),
    [
        (["1", "2"], [{"Title": "One"}], False, False, ValueError),
        (None, [{"Title": "One"}], False, False, TypeError),
        (["1"], None, False, False, TypeError),
        (["1"], ["not-a-dict"], False, False, TypeError),
        (["1"], [{"Title": "One"}], "yes", False, TypeError),
        (["1"], [{"Title": "One"}], False, "yes", TypeError),
    ],
)
def test_upload_input_validation(
    resolved_list: GbList,
    ids: object,
    rows: object,
    force: object,
    delete: object,
    error: type[Exception],
) -> None:
    with pytest.raises(error):
        resolved_list.upload(ids=ids, rows=rows, force=force, delete=delete)  # type: ignore[arg-type]
