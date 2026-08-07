from __future__ import annotations

from typing import Any

import pytest
import responses
from conftest import ACCESS_TOKEN, LIST_ID, LIST_NAME, SITE_ID

from graphbridge import GbList


def _cache_site(gb_list: GbList) -> None:
    gb_list.__dict__["_GbSite__site_data"] = {"id": SITE_ID}


def _items_url(gb_list: GbList, *, all_pages: bool = True) -> str:
    suffix = "/items?expand=fields&$top=200" if all_pages else "/items?expand=fields"
    return f"{gb_list.list_url}{suffix}"


def test_list_can_be_built_from_site_and_auth_keywords() -> None:
    gb_list = GbList(
        list_name=LIST_NAME,
        hostname="tenant.example.invalid",
        site_path="/sites/UnitTests",
        tenant_id="tenant-id",
        client_id="client-id",
        client_secret="client-secret",
    )

    assert gb_list.list_name == LIST_NAME
    assert gb_list.hostname == "tenant.example.invalid"


def test_list_rejects_invalid_site_object() -> None:
    with pytest.raises(TypeError, match="gb_site"):
        GbList(list_name=LIST_NAME, gb_site=object())


@pytest.mark.parametrize(("value", "error"), [("", ValueError), (123, TypeError)])
def test_list_name_validation(gb_list: GbList, value: object, error: type[Exception]) -> None:
    with pytest.raises(error):
        gb_list.list_name = value  # type: ignore[assignment]


def test_list_url_encodes_spaces(gb_list: GbList) -> None:
    _cache_site(gb_list)
    gb_list.list_name = "Project Tasks"

    assert gb_list.list_url == f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/Project%20Tasks"


@responses.activate
def test_list_data_is_requested_once_and_cached(gb_list: GbList, fixture_json: Any) -> None:
    _cache_site(gb_list)
    list_payload = fixture_json("list.json")
    responses.get(gb_list.list_url, json=list_payload, status=200)

    assert gb_list.list_data == list_payload
    assert gb_list.list_data is gb_list.list_data
    assert gb_list.list_id == LIST_ID
    assert len(responses.calls) == 1
    assert responses.calls[0].request.headers["Authorization"] == f"Bearer {ACCESS_TOKEN}"


@responses.activate
def test_list_data_error_raises_runtime_error(gb_list: GbList) -> None:
    _cache_site(gb_list)
    responses.get(gb_list.list_url, body="simulated missing list", status=404)

    with pytest.raises(RuntimeError, match="404 simulated missing list"):
        _ = gb_list.list_data


@responses.activate
def test_missing_list_id_returns_legacy_warning(gb_list: GbList) -> None:
    _cache_site(gb_list)
    responses.get(gb_list.list_url, json={}, status=200)

    assert gb_list.list_id == "<WARNING GbList | Element ID not found>"


@responses.activate
def test_list_items_all_follows_next_link(resolved_list: GbList, fixture_json: Any) -> None:
    first_page = fixture_json("items_page_1.json")
    second_page = fixture_json("items_page_2.json")
    responses.get(_items_url(resolved_list), json=first_page, status=200)
    responses.get(first_page["@odata.nextLink"], json=second_page, status=200)

    items = resolved_list.list_items_all

    assert [item["id"] for item in items] == ["1", "2", "3"]
    assert [call.request.method for call in responses.calls] == ["GET", "GET"]
    assert responses.calls[0].request.url == _items_url(resolved_list)
    assert responses.calls[1].request.url == first_page["@odata.nextLink"]
    assert all(call.request.headers["Authorization"] == f"Bearer {ACCESS_TOKEN}" for call in responses.calls)


@responses.activate
def test_list_items_all_error_raises_runtime_error(resolved_list: GbList) -> None:
    responses.get(_items_url(resolved_list), body="simulated server error", status=500)

    with pytest.raises(RuntimeError, match="500 simulated server error"):
        _ = resolved_list.list_items_all


@responses.activate
def test_list_items_returns_only_first_page(resolved_list: GbList, fixture_json: Any) -> None:
    first_page = fixture_json("items_page_1.json")
    responses.get(_items_url(resolved_list, all_pages=False), json=first_page, status=200)

    assert resolved_list.list_items == first_page["value"]
    assert len(responses.calls) == 1


@responses.activate
def test_list_items_error_raises_runtime_error(resolved_list: GbList) -> None:
    responses.get(_items_url(resolved_list, all_pages=False), body="simulated forbidden", status=403)

    with pytest.raises(RuntimeError, match="403 simulated forbidden"):
        _ = resolved_list.list_items


@responses.activate
def test_rows_and_ids_preserve_legacy_shapes(resolved_list: GbList) -> None:
    payload = {
        "value": [
            {"id": "1", "fields": {"Title": "One"}},
            {"id": "2", "fields": {"Title": "Two"}},
        ]
    }
    responses.get(_items_url(resolved_list), json=payload, status=200)
    responses.get(_items_url(resolved_list), json=payload, status=200)

    assert resolved_list.list_rows == [{"Title": "One"}, {"Title": "Two"}]
    assert resolved_list.list_ids == ["1", "2"]
    assert len(responses.calls) == 2


@responses.activate
def test_list_fields_performs_one_list_read(resolved_list: GbList) -> None:
    responses.get(
        _items_url(resolved_list),
        json={"value": [{"id": "1", "fields": {"Title": "One", "Status": "Open"}}]},
        status=200,
    )

    assert resolved_list.list_fields == ["Title", "Status"]
    assert len(responses.calls) == 1


@responses.activate
def test_list_fields_empty_list_performs_one_list_read(resolved_list: GbList) -> None:
    responses.get(_items_url(resolved_list), json={"value": []}, status=200)

    assert resolved_list.list_fields == []
    assert len(responses.calls) == 1


@responses.activate
def test_get_items_by_features_uses_or_and_nested_matching(resolved_list: GbList) -> None:
    payload = {
        "value": [
            {"id": "1", "fields": {"Status": "Open", "Category": "A"}},
            {"id": "2", "fields": {"Status": "Done", "Category": "B"}},
        ]
    }
    responses.get(_items_url(resolved_list), json=payload, status=200)
    responses.get(_items_url(resolved_list), json=payload, status=200)

    matched = resolved_list.get_items_by_features(
        [
            {"id": "1"},
            {"fields": {"Status": "Open", "Category": "A"}},
        ]
    )

    assert matched == [payload["value"][0]]
    # Legacy behavior: one complete fetch is performed for each feature dictionary.
    assert len(responses.calls) == 2


def test_get_items_by_features_with_no_features_returns_empty(resolved_list: GbList) -> None:
    assert resolved_list.get_items_by_features([]) == []


def test_field_name_encoding_and_decoding(resolved_list: GbList) -> None:
    row = {"Project 1": "Apollo", "Start-Date": "2026-01-01"}
    encoded = {
        "Project_x0020__x0031_": "Apollo",
        "Start_x002d_Date": "2026-01-01",
    }

    assert resolved_list.encode_map[" "] == "_x0020_"
    assert resolved_list.decode_map["_x0020_"] == " "
    assert resolved_list.encode_row(row) == encoded
    assert resolved_list.decode_row(encoded) == row


def test_list_repr_and_str_redact_secret_and_token(resolved_list: GbList) -> None:
    rendered = f"{resolved_list!r}\n{resolved_list!s}"

    assert resolved_list.client_secret not in rendered
    assert ACCESS_TOKEN not in rendered
    assert "<redacted>" in rendered
    assert LIST_ID in rendered
