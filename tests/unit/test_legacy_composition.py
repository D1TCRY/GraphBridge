from __future__ import annotations

import re
from types import SimpleNamespace

import responses
from conftest import ACCESS_TOKEN, LIST_ID, SITE_ID, FakeCredential

from graphbridge import GbAuth, GbList, GbSite, GraphBridgeClient
from graphbridge.legacy import GbAuth as LegacyGbAuth
from graphbridge.legacy import GbList as LegacyGbList
from graphbridge.legacy import GbSite as LegacyGbSite


class StaticCredential:
    def get_token(self, _scope: str) -> SimpleNamespace:
        return SimpleNamespace(token="new-api-token", expires_on=9999999999)


def test_legacy_module_exports_the_original_public_classes() -> None:
    assert (LegacyGbAuth, LegacyGbSite, LegacyGbList) == (GbAuth, GbSite, GbList)


def test_injected_legacy_objects_share_one_composed_client_and_session(
    auth: GbAuth,
    site: GbSite,
    gb_list: GbList,
) -> None:
    auth_client = auth._get_graph_client()

    assert site._get_graph_client() is auth_client
    assert gb_list._get_graph_client() is auth_client
    assert site._get_graph_client().transport.session is auth_client.transport.session
    assert gb_list._get_graph_client().transport.max_retries == 0
    assert len(FakeCredential.instances) == 1


@responses.activate
def test_legacy_adapter_renews_token_via_shared_credential(
    gb_list: GbList,
) -> None:
    responses.get(gb_list.site_url, json={"id": SITE_ID}, status=200)
    responses.get(gb_list.list_url, json={"id": LIST_ID}, status=200)

    assert gb_list.site_id == SITE_ID
    assert gb_list.list_id == LIST_ID

    credential = FakeCredential.instances[0]
    assert credential.get_token_calls == [
        "https://graph.microsoft.com/.default",
        "https://graph.microsoft.com/.default",
    ]
    assert all(call.request.headers["Authorization"] == f"Bearer {ACCESS_TOKEN}" for call in responses.calls)


@responses.activate
def test_legacy_and_new_api_expose_equivalent_raw_site_list_and_item_data(
    gb_list: GbList,
) -> None:
    site_payload = {"id": SITE_ID, "displayName": "Marketing"}
    list_payload = {"id": LIST_ID, "displayName": "Tasks"}
    item_payload = {"value": [{"id": "1", "fields": {"Title": "One"}}]}

    responses.get(gb_list.site_url, json=site_payload, status=200)
    responses.get(gb_list.list_url, json=list_payload, status=200)
    responses.get(f"{gb_list.list_url}/items?expand=fields&$top=200", json=item_payload, status=200)

    legacy_site_data = gb_list.site_data
    legacy_list_data = gb_list.list_data
    legacy_items = gb_list.list_items_all

    client = GraphBridgeClient(credential=StaticCredential(), max_retries=0)
    responses.get(gb_list.site_url, json=site_payload, status=200)
    responses.get(gb_list.list_url, json=list_payload, status=200)
    responses.get(re.compile(r"https://graph\.microsoft\.com/v1\.0/.*/items.*"), json=item_payload, status=200)

    site = client.sites.get_by_path(hostname=gb_list.hostname, path=gb_list.site_path)
    tasks = site.lists.get(gb_list.list_name)
    new_items = list(tasks.items.iter_all(top=200))

    assert site.info.raw == legacy_site_data
    assert tasks.info.raw == legacy_list_data
    assert [item.raw for item in new_items] == legacy_items


def test_legacy_secret_token_and_composed_client_are_not_rendered(auth: GbAuth) -> None:
    _ = auth.token
    client = auth._get_graph_client()

    rendered = f"{auth!r} {auth!s} {client!r} {client.transport!r}"
    assert auth.client_secret not in rendered
    assert ACCESS_TOKEN not in rendered
