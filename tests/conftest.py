from __future__ import annotations

import json
import os
import socket
from collections.abc import Iterator
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest
from azure.identity import ClientSecretCredential

import graphbridge.graph_bridge as graph_bridge_module
from graphbridge import GbAuth, GbList, GbSite, GraphBridgeClient
from graphbridge.resources import SharePointListResource


TENANT_ID = "00000000-0000-0000-0000-000000000001"
CLIENT_ID = "00000000-0000-0000-0000-000000000002"
CLIENT_SECRET = "unit-test-client-secret"
ACCESS_TOKEN = "unit-test-access-token"
HOSTNAME = "tenant.example.invalid"
SITE_PATH = "/sites/UnitTests"
LIST_NAME = "Tasks"
SITE_ID = "example.invalid,site-collection-id,site-web-id"
LIST_ID = "00000000-0000-0000-0000-000000000003"
INTEGRATION_CONFIRMATION = "GRAPHBRIDGE_DEDICATED_TEST_ENVIRONMENT"
INTEGRATION_WRITE_CONFIRMATION = "CREATE_UPDATE_DELETE_OWN_ITEMS"
INTEGRATION_DEDICATED_PREFIX = "GraphBridge Integration - "


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption(
        "--run-integration",
        action="store_true",
        default=False,
        help="run opt-in tests against the explicitly configured dedicated test site",
    )


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    if config.getoption("--run-integration"):
        return
    skipped = pytest.mark.skip(reason="integration tests require --run-integration and dedicated environment variables")
    for item in items:
        if "integration" in item.keywords:
            item.add_marker(skipped)


def _integration_environment() -> dict[str, str]:
    names = (
        "AZURE_TENANT_ID",
        "AZURE_CLIENT_ID",
        "AZURE_CLIENT_SECRET",
        "GRAPHBRIDGE_INTEGRATION_SITE_ID",
        "GRAPHBRIDGE_INTEGRATION_SITE_NAME",
        "GRAPHBRIDGE_INTEGRATION_LIST_ID",
        "GRAPHBRIDGE_INTEGRATION_LIST_NAME",
        "GRAPHBRIDGE_INTEGRATION_CONFIRM",
    )
    values = {name: os.environ.get(name, "").strip() for name in names}
    if any(not value for value in values.values()):
        pytest.skip("dedicated integration configuration is incomplete")
    if values["GRAPHBRIDGE_INTEGRATION_CONFIRM"] != INTEGRATION_CONFIRMATION:
        pytest.skip("dedicated integration confirmation is absent")
    for name in ("GRAPHBRIDGE_INTEGRATION_SITE_NAME", "GRAPHBRIDGE_INTEGRATION_LIST_NAME"):
        if not values[name].startswith(INTEGRATION_DEDICATED_PREFIX):
            pytest.fail(
                f"{name} must start with {INTEGRATION_DEDICATED_PREFIX!r}",
                pytrace=False,
            )
    return values


@pytest.fixture(scope="session")
def integration_client() -> Iterator[GraphBridgeClient]:
    config = _integration_environment()
    credential = ClientSecretCredential(
        tenant_id=config["AZURE_TENANT_ID"],
        client_id=config["AZURE_CLIENT_ID"],
        client_secret=config["AZURE_CLIENT_SECRET"],
    )
    with GraphBridgeClient(credential=credential) as client:
        yield client


@pytest.fixture(scope="session")
def dedicated_list(integration_client: GraphBridgeClient) -> SharePointListResource:
    config = _integration_environment()
    site = integration_client.sites.get(config["GRAPHBRIDGE_INTEGRATION_SITE_ID"])
    if site.info.display_name != config["GRAPHBRIDGE_INTEGRATION_SITE_NAME"]:
        pytest.fail("configured site ID does not match the dedicated site name", pytrace=False)
    sharepoint_list = site.lists.get_by_id(config["GRAPHBRIDGE_INTEGRATION_LIST_ID"])
    if sharepoint_list.display_name != config["GRAPHBRIDGE_INTEGRATION_LIST_NAME"]:
        pytest.fail("configured list ID does not match the dedicated list name", pytrace=False)
    return sharepoint_list


@pytest.fixture
def integration_writes_enabled() -> None:
    if os.environ.get("GRAPHBRIDGE_INTEGRATION_ALLOW_WRITES") != INTEGRATION_WRITE_CONFIRMATION:
        pytest.skip("write integration test requires an additional explicit confirmation")


class FakeCredential:
    instances: list[FakeCredential] = []

    def __init__(self, tenant_id: str, client_id: str, client_secret: str) -> None:
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.get_token_calls: list[str] = []
        type(self).instances.append(self)

    def get_token(self, scope: str) -> SimpleNamespace:
        self.get_token_calls.append(scope)
        return SimpleNamespace(token=ACCESS_TOKEN)


@pytest.fixture(autouse=True)
def fake_azure_credential(monkeypatch: pytest.MonkeyPatch) -> None:
    FakeCredential.instances.clear()
    monkeypatch.setattr(graph_bridge_module, "ClientSecretCredential", FakeCredential)


@pytest.fixture(autouse=True)
def block_real_network(monkeypatch: pytest.MonkeyPatch, request: pytest.FixtureRequest) -> None:
    if request.node.get_closest_marker("integration") is not None:
        return

    def blocked_connect(*_args: object, **_kwargs: object) -> None:
        raise AssertionError("Unit tests must not open real network connections")

    monkeypatch.setattr(socket.socket, "connect", blocked_connect)


@pytest.fixture
def auth() -> GbAuth:
    return GbAuth(TENANT_ID, CLIENT_ID, CLIENT_SECRET)


@pytest.fixture
def site(auth: GbAuth) -> GbSite:
    return GbSite(hostname=HOSTNAME, site_path=SITE_PATH, gb_auth=auth)


@pytest.fixture
def gb_list(site: GbSite) -> GbList:
    return GbList(list_name=LIST_NAME, gb_site=site)


@pytest.fixture
def resolved_list(gb_list: GbList) -> GbList:
    gb_list.__dict__["_GbSite__site_data"] = {"id": SITE_ID}
    gb_list.__dict__["_GbList__list_data"] = {"id": LIST_ID}
    return gb_list


@pytest.fixture
def fixture_json() -> Any:
    fixture_dir = Path(__file__).parent / "fixtures"

    def load(name: str) -> dict[str, Any]:
        return json.loads((fixture_dir / name).read_text(encoding="utf-8"))

    return load
